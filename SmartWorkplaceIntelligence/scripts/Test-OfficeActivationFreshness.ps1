# Version 1.0.0
[CmdletBinding()]
param([string]$TestRoot = (Join-Path ([IO.Path]::GetTempPath()) 'SmartM365-OfflineTests'))
# Synthetic local files only. No collectors, tenant calls, production data or uploads.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:checks = 0
function Check([bool]$Pass, [string]$Message) { if (-not $Pass) { throw $Message }; $script:checks++ }
function Import-ScriptFunction([string]$File, [string]$Name) {
    $tokens = $null; $errors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot $File), [ref]$tokens, [ref]$errors)
    if ($errors.Count) { throw ($errors.Message -join '; ') }
    $node = $ast.FindAll({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] }, $true) | Where-Object Name -eq $Name
    if (@($node).Count -ne 1) { throw "Function lookup failed: $Name" }
    [scriptblock]::Create($node.Extent.Text)
}

$now = [datetime]::new(2026, 10, 22, 12, 0, 0, [DateTimeKind]::Utc)
$fixture = Join-Path ([IO.Path]::GetFullPath($TestRoot)) ([guid]::NewGuid().ToString('N'))
$raw = Join-Path $fixture 'raw'
$last = Join-Path $raw 'DATA-LAST'
New-Item -ItemType Directory -Path $last -Force | Out-Null
$activationPath = Join-Path $last 'M365_Apps_Activations.csv'
$trustPath = Join-Path $fixture 'DataTrustEvidence.csv'

function Write-ActivationFixture([datetime]$ReportDate, [datetime]$AcquiredUtc) {
    [pscustomobject][ordered]@{
        TenantKey = 'synthetic-test'; RunId = 'run-1'; 'Report Refresh Date' = $ReportDate.ToString('yyyy-MM-dd')
        'User Principal Name' = 'user1@example.test'; 'Product Type' = 'MICROSOFT 365 APPS FOR ENTERPRISE'
        'Last Activated Date' = $ReportDate.AddDays(-3).ToString('yyyy-MM-dd'); Windows = '1'; Mac = '0'
    } | Export-Csv -LiteralPath $activationPath -NoTypeInformation
    [IO.File]::SetLastWriteTimeUtc($activationPath, $AcquiredUtc)
}
function Get-ActivationTrust([int]$WarningAgeHours = 0) {
    $trustArgs = @{DataRoot = $raw; OutputPath = $trustPath; AsOfDate = $now.Date; AsOfUtc = $now}
    if ($WarningAgeHours) { $trustArgs.OfficeActivationWarningAgeHours = $WarningAgeHours }
    & (Join-Path $PSScriptRoot 'New-DataTrustEvidence.ps1') @trustArgs | Out-Null
    $rows = @(Import-Csv -LiteralPath $trustPath | Where-Object 'File Name' -eq 'M365_Apps_Activations.csv')
    Check ($rows.Count -eq 1) 'Data Trust must carry exactly one Office activation source row'
    $rows[0]
}

try {
    # Fresh: acquired two hours ago; Graph's usual two-day report lag must not warn.
    Write-ActivationFixture $now.AddDays(-2) $now.AddHours(-2)
    $fresh = Get-ActivationTrust
    Check ($fresh.'Evidence Domain' -eq 'Applications' -and $fresh.'Source Name' -eq 'Microsoft 365 Apps activations') 'Office activation source identity changed'
    Check ($fresh.'Freshness State' -eq 'Fresh' -and $fresh.'Decision Readiness' -eq 'Ready' -and $fresh.'Evidence Status' -eq 'Observed') 'Fresh acquisition was not reported Fresh'
    Check ($fresh.'Snapshot Date' -eq $now.AddDays(-2).ToString('yyyy-MM-dd')) 'Snapshot date must stay the business Report Refresh Date'
    Check ($fresh.'Snapshot Basis' -match '^Business snapshot: Report Refresh Date; last acquisition 2026-10-22T10:00:00Z, Aging after 72 hours$') "Acquisition basis not traced: $($fresh.'Snapshot Basis')"

    # Kept old: the producer left the previous export (written 100 hours ago) unchanged.
    Write-ActivationFixture $now.AddDays(-6) $now.AddHours(-100)
    $kept = Get-ActivationTrust
    Check ($kept.'Freshness State' -eq 'Aging') 'Kept export past the threshold was not reported Aging'
    Check ($kept.'Decision Readiness' -eq 'Ready' -and $kept.'Evidence Status' -eq 'Observed') 'Aging warning must never block the source'
    Check ($kept.'Recommended Action' -match 'source producer') 'Aging source must point at the producer'
    $tolerant = Get-ActivationTrust -WarningAgeHours 120
    Check ($tolerant.'Freshness State' -eq 'Fresh' -and $tolerant.'Snapshot Basis' -match 'Aging after 120 hours') 'Configured threshold was ignored'

    # A business date already Stale stays Stale and still requires a refresh.
    Write-ActivationFixture $now.AddDays(-40) $now.AddHours(-100)
    $stale = Get-ActivationTrust
    Check ($stale.'Freshness State' -eq 'Stale' -and $stale.'Decision Readiness' -eq 'Refresh required') 'Stale business date was softened by the acquisition rule'

    # Absent: Data Trust reports it missing; Workforce keeps requiring it.
    Remove-Item -LiteralPath $activationPath
    $missing = Get-ActivationTrust
    Check ($missing.'File State' -eq 'Missing' -and $missing.'Decision Readiness' -eq 'Blocked' -and $missing.'Freshness State' -eq 'Unknown') 'Missing Office activation export was not reported Blocked'

    . (Import-ScriptFunction 'New-WorkforceIdentityEvidence.ps1' 'Get-OfficeActivationFreshnessWarning')
    $reference = [datetimeoffset]$now
    Check ($null -eq (Get-OfficeActivationFreshnessWarning -AcquiredUtc $now.AddHours(-2) -LatestReportDate $now.AddDays(-2) -ReferenceUtc $reference -WarningAgeHours 72)) 'Fresh acquisition raised a warning'
    Check ($null -eq (Get-OfficeActivationFreshnessWarning -AcquiredUtc $now.AddHours(-72) -LatestReportDate $now.AddDays(-5) -ReferenceUtc $reference -WarningAgeHours 72)) 'Warning raised at the threshold instead of past it'
    $warning = Get-OfficeActivationFreshnessWarning -AcquiredUtc $now.AddHours(-100) -LatestReportDate $now.AddDays(-6).Date -ReferenceUtc $reference -WarningAgeHours 72
    Check ($warning -match '^Office activation report is aging: M365_Apps_Activations\.csv was last acquired 2026-10-18T08:00:00Z \(100 hours ago; warning after 72 hours\); latest Report Refresh Date=2026-10-16\.') "Kept export warning text changed: $warning"
    $unknown = Get-OfficeActivationFreshnessWarning -AcquiredUtc $now.AddHours(-100) -LatestReportDate $null -ReferenceUtc $reference -WarningAgeHours 72
    Check ($unknown -match 'latest Report Refresh Date=unknown\.') 'Missing report date was not reported unknown'

    # Wiring: Workforce warns through the relayed warning stream and still requires the file.
    $workforce = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'New-WorkforceIdentityEvidence.ps1') -Raw
    Check ($workforce.Contains("Get-RequiredCsv 'DATA-LAST\M365_Apps_Activations.csv'")) 'Workforce no longer requires the Office activation export'
    Check ($workforce.Contains('if ($officeActivationWarning) { Write-Warning $officeActivationWarning }')) 'Workforce does not emit the freshness warning'
    $build = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'Invoke-PreparedEvidenceBuild.ps1') -Raw
    Check ($build.Contains("`$wf+=@('-OfficeActivationWarningAgeHours',[string]`$OfficeActivationWarningAgeHours)") -and $build.Contains("`$trustArgs+=@('-OfficeActivationWarningAgeHours',[string]`$OfficeActivationWarningAgeHours)")) 'Build does not pass the threshold to Workforce and Data Trust'
    $pipeline = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'PreparedEvidencePipeline.psm1') -Raw
    Check ($pipeline.Contains("`$systemPolicyArgs+=@('-OfficeActivationWarningAgeHours',`$OfficeActivationWarningAgeHours)") -and $pipeline.Contains('$provenance.OfficeActivationFreshness=')) 'Pipeline does not pass or record the threshold'
    $prepareRoot = Join-Path $PSScriptRoot '../../SmartM365/SmartInventory/PreparedEvidence'
    $prepare = Get-Content -LiteralPath (Join-Path $prepareRoot 'SmartM365-WorkplaceEvidence-Prepare.ps1') -Raw
    Check ($prepare.Contains('OfficeActivationWarningAgeHours=$officeActivationWarning')) 'Prepare does not pass the threshold setting'
    # Template keys are written into the private configuration, whose fingerprint retention compares across runs.
    $template = Get-Content -LiteralPath (Join-Path $prepareRoot 'SmartM365-WorkplaceEvidence-Prepare.local.json.template') -Raw
    Check (-not $template.Contains('PreparedOfficeActivationWarningAgeHours')) 'The optional threshold must stay out of the template'
    $resolution = [regex]::Match($prepare, "(?ms)^\s*\`$officeActivationWarning=ConfigValue 'PreparedOfficeActivationWarningAgeHours'\r?\n.*?must be between 1 and 8760\.'\}\r?\n")
    Check $resolution.Success 'Prepare threshold resolution not found'
    $resolve = [scriptblock]::Create('param($Configured) function ConfigValue([string]$Name) { $Configured }' + "`n" + $resolution.Value + '$officeActivationWarning')
    Check ((& $resolve $null) -eq 72 -and (& $resolve '') -eq 72) 'Absent setting must default to 72 hours'
    Check ((& $resolve 120) -eq 120) 'Configured setting was ignored'
    $rejected = $false; try { & $resolve 0 | Out-Null } catch { $rejected = $true }
    Check $rejected 'Out-of-range setting was accepted'
    foreach ($worker in 'New-WorkforceIdentityEvidence.ps1', 'New-DataTrustEvidence.ps1', 'Invoke-PreparedEvidenceBuild.ps1') {
        $tokens = $null; $errors = $null
        $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot $worker), [ref]$tokens, [ref]$errors)
        $parameter = @($ast.ParamBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq 'OfficeActivationWarningAgeHours' })
        Check ($parameter.Count -eq 1 -and $parameter[0].DefaultValue.Extent.Text -eq '72') "Default threshold must be 72 hours: $worker"
    }
}
finally {
    Remove-Item -LiteralPath $fixture -Recurse -Force -ErrorAction SilentlyContinue
}
Write-Host "PASS: $script:checks Office activation freshness checks. Synthetic local evidence only; no collectors or live writes."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCoP8Iyho5iwwiE
# bxe73lcFUwK7zOtYZl9pt1/Ai/rpyqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEID+G3o6KMJnmq7OkFjQ2qQr2PY0Udisab72JFvUk1bMvMA0GCSqG
# SIb3DQEBAQUABIIBgHaRkdHEsXJEx4oy8wwFEbdHr6bqgpuFhUR9nOmx1quP+ZkJ
# i0J8bDWC+cwQAImzuYf1lPsoeqa2+ghO6hPlGoo5OuLHL9UEnARo4jrGfkW/VM+b
# oX/QuLjcLkgimX5U5sv+F5yHwDlQrjqkmolr27sekcLgqyU4EtYmSpHaBDLlvAVu
# +CYs0bP5z/wkI7IUegeufr1rgxGA+YVvvEh8JLQbVVjLvoIH2NvSEkSg70OvZiJo
# 1++qlTO76/eIeibR3VQ62N/uwhNaAAOqjKXNOFJ3Qp7xmJSEqKPE5m/2Fxc8aYhP
# 0p8Wq9kL1lc2gB5HxaU1DUlAZyzcCxZZwJwSVFFgnJ2/nawfbIUMxi9yuWIYd6vl
# NN0ZDiAYNtbUAYsH2QzsP5JjyGqzrGs7hH4vBCBdVbj7duKGJWFneGgDXXTx/ktN
# KzF60jOHdfSibTBVxhI7zWoJA3B/WCmhjWhSwlyzD/uK7+8WngeRmudE7ipr5x49
# vW4rcNG56pQ5QuZTpaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjIxNDM0
# NTBaMC8GCSqGSIb3DQEJBDEiBCDziFwGtpOVBAGOyhhLP4+a+BBREfsIvkkXKeNE
# Ov7LizANBgkqhkiG9w0BAQEFAASCAgChGRUwKBlqyTgWAL5Z+0S0SIdGv/UZz6un
# V4qMbaRJzc9KOONBMOBbCPyFmIQT1I1kuZIb76aVWkSpOpyw29iPLtGfE0nYtT2M
# lCQ74KjkkUvQTWSFKYsqoTUm6FQk6WI4AE1VwLV8z3EBBnPr8FG5DZc1p7r85t3q
# bEirHxVOkGchQCEpIz4ipbyVPRTH4wi8BZgbtf0yJz/bUVnKYN/rZThzDKtga7Uc
# 9DgF92+Rrif77y3eFXkyJIF+hDxVt90urLkDiRZ9Ef5bBbps2BtXIn3rRZDC8j7w
# OcG1G4giGTKzVolyY2kYVLWon0EqcmDW/CAz0dk4ScB1AQRSVq9+U6RhxELparbQ
# HgKrMzZhcuCwdghy7dfBqbsw6m3rsEae/E5WbosvqP1Ix5o57jO/qsf+I+vl+BjT
# 5hpcgxeHLziNpueTxra4AqxBIKyhYk2k/P8RELDlhicJCTcwaLdCpw/KhMPqgICO
# ubJ3QdgDuXW2KbyVpRmA5oDhqeBAR2cmWbutUX4oqr1+e2NzAuTD2UUdWh4TeRTj
# 1j91/7EXGMZbW6pDGsDvli5Gmrw6/Zu74RTAQzOPd4Er0gaK7MkUnEAHy7ilvnwx
# 07SlOP5FzSd2EPkHtcoiL44n3ZOfnuPl/WfRiIlU1FEL+OCOBibemFuhvzaxW2Uf
# ypuWA5LXHw==
# SIG # End signature block
