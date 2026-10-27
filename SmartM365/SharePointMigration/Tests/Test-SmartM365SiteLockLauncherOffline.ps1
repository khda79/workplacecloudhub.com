<#
.SYNOPSIS
    Exercises site lock selection and the SharePoint Server scope guard offline.
#>
#Requires -Version 5.1
$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSEdition -ne 'Desktop') { throw 'Run this test with Windows PowerShell 5.1.' }

$toolkit = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$fixture = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-SiteLockTest-' + [guid]::NewGuid().ToString('N'))
$migration = Join-Path $fixture 'Migrations\Synthetic'
$operations = Join-Path $fixture 'Scripts\Operations'
$launchers = Join-Path $fixture 'Scripts\Launchers'
$generic = Join-Path $launchers 'Generic'

function Assert-Test {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}

try {
    foreach ($directory in @($migration, $operations, $generic)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }
    @'
@{
    Name = 'Synthetic'
    Source = @{ Type = 'SP2019'; WebApplicationUrl = 'https://farm.example' }
    Target = @{ Type = 'SPO'; SiteUrl = 'https://tenant.sharepoint.com/sites/destination'; TenantAdminUrl = 'https://tenant-admin.sharepoint.com' }
}
'@ | Set-Content -LiteralPath (Join-Path $migration 'migration.config.psd1') -Encoding UTF8
    'https://farm.example/sites/source/sub https://tenant.sharepoint.com/sites/destination' |
        Set-Content -LiteralPath (Join-Path $migration 'migration.mapping.txt') -Encoding UTF8
    Copy-Item -LiteralPath (Join-Path $toolkit 'Scripts\Launchers\Generic\SmartM365-SharePointMigration-SiteLockStateLauncher.ps1') -Destination $generic
    Copy-Item -LiteralPath (Join-Path $toolkit 'Scripts\Launchers\SmartM365-SharePointMigration-ConsoleLifecycle.ps1') -Destination $launchers
    Copy-Item -LiteralPath (Join-Path $toolkit 'Scripts\Launchers\SmartM365-SharePointMigration-Storage.ps1') -Destination $launchers
    Copy-Item -LiteralPath (Join-Path $toolkit 'Scripts\Operations\SmartM365-SharePointSource-SiteLockStateSet.ps1') -Destination $operations
    Copy-Item -LiteralPath (Join-Path $toolkit 'Scripts\Operations\SmartM365-SharePointTarget-SiteLockStateSet.ps1') -Destination $operations
    $launcher = Join-Path $generic 'SmartM365-SharePointMigration-SiteLockStateLauncher.ps1'
    $sourceOperation = Join-Path $operations 'SmartM365-SharePointSource-SiteLockStateSet.ps1'

    $preview = (& $launcher -ToolkitRoot $fixture -MigrationName Synthetic -Side Source -LockState ReadOnly -PreviewOnly 6>&1 | Out-String)
    Assert-Test ($preview -match 'Source \(SP2019\)' -and $preview -match 'PreviewOnly completed') 'Source preview did not show the selected side and mode.'
    Assert-Test (-not (Test-Path (Join-Path $operations 'Output'))) 'Preview unexpectedly ran an operation.'

    $global:SiteLockTestAnswers = New-Object 'System.Collections.Generic.Queue[string]'
    @('1', '2', '3') | ForEach-Object { $global:SiteLockTestAnswers.Enqueue($_) }
    function global:Read-Host { param([string]$Prompt) return $global:SiteLockTestAnswers.Dequeue() }
    $menu = (& $launcher -ToolkitRoot $fixture -PreviewOnly 6>&1 | Out-String)
    Assert-Test ($menu -match 'Destination \(SPO\)' -and $menu -match 'Requested lock state: NoAccess') 'Interactive menu selected the wrong site or mode.'

    $global:SiteLockTestAnswers = New-Object 'System.Collections.Generic.Queue[string]'
    $global:SiteLockTestAnswers.Enqueue('NO')
    $cancelled = (& $launcher -ToolkitRoot $fixture -MigrationName Synthetic -Side Destination -LockState Unlock 6>&1 | Out-String)
    Assert-Test ($cancelled -match 'Cancelled; no site change was made') 'Final confirmation did not cancel.'
    Assert-Test (-not (Test-Path (Join-Path $operations 'Output'))) 'Cancellation unexpectedly ran an operation.'

    $targetOperation = Join-Path $operations 'SmartM365-SharePointTarget-SiteLockStateSet.ps1'
    @'
param([string]$SiteUrl, [string]$TenantAdminUrl, [string]$LockState)
"$SiteUrl|$TenantAdminUrl|$LockState" | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'target-call.txt')
'@ | Set-Content -LiteralPath $targetOperation -Encoding UTF8
    $global:SiteLockTestAnswers = New-Object 'System.Collections.Generic.Queue[string]'
    $global:SiteLockTestAnswers.Enqueue('EXECUTE')
    & $launcher -ToolkitRoot $fixture -MigrationName Synthetic -Side Destination -LockState Unlock 6>&1 | Out-Null
    $targetCall = Get-Content -LiteralPath (Join-Path $operations 'target-call.txt') -Raw
    Assert-Test ($targetCall.Trim() -eq 'https://tenant.sharepoint.com/sites/destination|https://tenant-admin.sharepoint.com|Unlock') 'Destination dispatch passed the wrong URL, admin URL, or state.'
    Remove-Item function:Read-Host -ErrorAction SilentlyContinue

    $global:SiteLockTestUrl = 'https://farm.example/sites/source'
    $global:SiteLockTestState = 'Unlock'
    $global:SiteLockTestAdminState = 'Unlock'
    $global:SiteLockTestSetCalls = 0
    function global:Get-SPSite {
        param([string]$Identity)
        return [pscustomobject]@{
            Url = $global:SiteLockTestUrl
            ReadOnly = ($global:SiteLockTestState -eq 'ReadOnly')
        }
    }
    function global:Get-SPSiteAdministration {
        param($Identity)
        if ($Identity -isnot [string] -or $Identity -cne $global:SiteLockTestUrl) {
            throw 'Site administration lookup requires the exact site collection URL as a string.'
        }
        return [pscustomobject]@{
            ReadLocked = ($global:SiteLockTestAdminState -eq 'NoAccess')
            ReadOnly = ($global:SiteLockTestAdminState -eq 'ReadOnly')
            WriteLocked = ($global:SiteLockTestAdminState -eq 'NoAdditions')
        }
    }
    function global:Set-SPSite {
        param([string]$Identity, [string]$LockState, [switch]$Confirm)
        if ($Identity -cne $global:SiteLockTestUrl) { throw 'Set-SPSite used the wrong site collection URL.' }
        $global:SiteLockTestSetCalls++
        $global:SiteLockTestState = $LockState
    }
    $blocked = $false
    try { & $sourceOperation -SiteUrl 'https://farm.example/sites/source/sub' -LockState ReadOnly 2>&1 | Out-Null }
    catch { $blocked = $_.Exception.Message -like '*not the exact site collection URL*' }
    Assert-Test $blocked 'A subsite URL was not rejected.'
    Assert-Test ($global:SiteLockTestSetCalls -eq 0) 'Set-SPSite ran for a subsite URL.'

    & $sourceOperation -SiteUrl $global:SiteLockTestUrl -LockState ReadOnly -WhatIfMode 6>&1 | Out-Null
    Assert-Test ($global:SiteLockTestSetCalls -eq 0) 'WhatIfMode changed the mocked site.'
    & $sourceOperation -SiteUrl $global:SiteLockTestUrl -LockState ReadOnly 6>&1 | Out-Null
    Assert-Test ($global:SiteLockTestSetCalls -eq 1 -and $global:SiteLockTestState -eq 'ReadOnly') 'Exact site collection did not pass the mocked operation and read-back.'
    Assert-Test ($global:SiteLockTestAdminState -eq 'Unlock') 'The administration mock did not preserve the stale state.'
    & $sourceOperation -SiteUrl $global:SiteLockTestUrl -LockState ReadOnly 6>&1 | Out-Null
    Assert-Test ($global:SiteLockTestSetCalls -eq 1) 'An already read-only site was changed again.'
    $reports = @(Get-ChildItem -LiteralPath (Join-Path $operations 'Output') -Filter '*.csv')
    Assert-Test ($reports.Count -eq 4) 'Expected one CSV for each source operation attempt.'
    Assert-Test (@($reports | ForEach-Object { (Import-Csv -LiteralPath $_.FullName -Delimiter ';').Status }) -contains 'Success') 'Successful read-back was not recorded.'
    Assert-Test (@($reports | ForEach-Object { (Import-Csv -LiteralPath $_.FullName -Delimiter ';').Status }) -contains 'AlreadySet') 'An already read-only site was not recognized.'
    Assert-Test (@(Get-ChildItem -LiteralPath (Join-Path $operations 'logs') -Filter '*.log' |
        Select-String -Pattern 'AdminReadOnly=False; AdminWriteLocked=False; SiteReadOnly=True').Count -gt 0) 'Divergent lock indicators were not logged.'

    Write-Host 'Site lock launcher offline tests passed.'
}
finally {
    Remove-Item function:Read-Host -ErrorAction SilentlyContinue
    Remove-Item function:Get-SPSite -ErrorAction SilentlyContinue
    Remove-Item function:Get-SPSiteAdministration -ErrorAction SilentlyContinue
    Remove-Item function:Set-SPSite -ErrorAction SilentlyContinue
    Remove-Variable SiteLockTestAnswers,SiteLockTestUrl,SiteLockTestState,SiteLockTestAdminState,SiteLockTestSetCalls -Scope Global -ErrorAction SilentlyContinue
    $resolved = [IO.Path]::GetFullPath($fixture)
    $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if ($resolved.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -and
        [IO.Path]::GetFileName($resolved).StartsWith('SmartM365-SiteLockTest-', [StringComparison]::Ordinal)) {
        Remove-Item -LiteralPath $resolved -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAkjO80KkGP+qKe
# k/Xo9W5JHHJVGm9Al6JhbXjy2H95yKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIENOKewvkqtEwULb0JkOVez/1HL0TGRt68BKlH9VZEmGMA0GCSqG
# SIb3DQEBAQUABIIBgItMAKlSXNi+yDZygC8UCX7tXgvzKA06UUVEBaT+bJNqyM5x
# ELIBHEvssHO24Ax01azOLSUiIFrHUluVFwi/99qi7GlUEvzXCY7mSom3Zd3SAIrt
# Iz1wgu/6HdDyEqkRQ0w+Wq8+Ipavxt9meygeCH9VI8G+eAHTV5NtZQgC+xC/3P8a
# OFxpkTUBKWmdIsuJQtiOyCbJi/Th8wwHszhfPY/YeAmHqX673q1eH4W7TEkNnwUf
# J6VAEaF+3AW0vWmgL2wLhQp9Et70DpaBgwO80X0y3xYero+i/2TSnoQobJJIn/U7
# h5Q9oMlsx2mgw6CMVbxphersKIBZrBrFRAQe5qPaNQh6a27XXNmY7rrxKyvOYNBu
# mEzNSoueYAGmaQTjHUV4QDHGC4bF/wWCQDrKDbRBs7i+3xGfnEZPfD42sdXKxg8O
# 4lC0UE4FowFpHTtpLmknrNIGYtzb8y/PQ7ui9RKHDi3yCJxQ5ARKkI0zWN2BWyqq
# VM9gcWnNb3O4rKaxlaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTExMDI0
# NTBaMC8GCSqGSIb3DQEJBDEiBCCiMwHWUEXZVChkh8qpXx6VSU6VH7LcryCXJvM3
# Y/UjZjANBgkqhkiG9w0BAQEFAASCAgBTof/8Mtgh5YblLnBoIJk6QmSP/Mhb9YST
# ckRzt+FbaNDYsZPLravzzcDRLfjCchLhB7LMFcUCaB0l/SgUVZ1Qt0B9ZIeDD/nm
# xKfoME5/5N9tNF0Ph/s0ex3PX2tG0tk8L7JoEcMKnh5V3unsdpZexFAh5o1+6pHR
# yDqdikMG93y65WLooP/LD2LtHgKg3bR30ymLHRFw7TdJSysQN+CJj5TJlsManp7v
# zGK7pHU9CeN/UzG0wjoQkCSZPhxXkbEdoiIqsbvpOZ43/DqV5UyotdW22768Y3qv
# y4TatgkUkQHXzuD4N/2QlfpqQ2Pc0syB8bRMb/YAsyLsZtbsOnrY3DAk2lPFDyxU
# x3z9JlU5Ye/mXdmpSWUE9RYAJlkxE2LiLCMKIlZkslLsAlM8E2wyPPxALES/+bu8
# V+SnmWuNSKVLyWRlX/jVyIJndvZDtULJ9ON/qrnDRs6DraR2XV5MLne2wgz+c//C
# gnJPTWCrd+6w6Ydiaf1XNCoUOg0e2JfASoeyUwmX0rXDB55R18WkD07JgHMbUkMp
# lGL6L2PL35EJ3vTw0CsVcDhdJK1dvqKdQCzjvqiNc/OO5S2Q99TmJVC4hJh70YvF
# l4ezyihmn9+qcLIBA2HdGs31e4FO5bd0O2ztKH3gXCehZFyhyAftDZwSz3EBS5K3
# A3E8z9jiyg==
# SIG # End signature block
