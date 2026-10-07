[CmdletBinding()]
param([string]$TestRoot = (Join-Path ([IO.Path]::GetTempPath()) 'SmartM365-OfflineTests'))
# Synthetic local files only. No collectors, tenant calls, production data or uploads.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$fixture = Join-Path ([IO.Path]::GetFullPath($TestRoot)) ([guid]::NewGuid().ToString('N'))
$raw = Join-Path $fixture 'raw'
$last = Join-Path $raw 'DATA-LAST'
New-Item -ItemType Directory -Path $last,(Join-Path $fixture 'config') -Force | Out-Null
$adPath = Join-Path $last 'AD_HealthCheck.csv'
$requiredPath = Join-Path $last 'Required.csv'
$contract = Join-Path $fixture 'config/prepared-source-contract.json.txt'
@{currentFiles=@('Required.csv');optionalCurrentFiles=@(@{file='AD_HealthCheck.csv';maxAgeHours=48});mappingFiles=@();dailyFiles=@();history=@()} |
    ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $contract
[pscustomobject]@{TenantKey='synthetic-test';Value=1} | Export-Csv -LiteralPath $requiredPath -NoTypeInformation
$module = Import-Module (Join-Path $PSScriptRoot 'PreparedEvidencePipeline.psm1') -Force -PassThru
Import-Module (Join-Path $PSScriptRoot 'OptionalMonitoringEvidence.psm1') -Force
& $module {param($root) $script:ProductRoot=$root} $fixture
$script:checks=0
function Check([bool]$Pass,[string]$Message) { if (-not $Pass) { throw $Message }; $script:checks++ }
function Reject([scriptblock]$Action,[string]$Pattern) {
    $message=''; try { & $Action | Out-Null } catch { $message=$_.Exception.Message }
    Check ($message -match $Pattern) "Expected '$Pattern', got '$message'"
}
$now = [datetime]::UtcNow
function Write-ADFixture([double]$AgeHours=1,[string[]]$Statuses=@('OK','Warning','Critical','NotMeasured'),[string]$Tenant='synthetic-test') {
    @($Statuses | ForEach-Object {
        [pscustomobject]@{TenantKey=$Tenant;Forest='example.invalid';Domain='example.invalid';Category='Replication';Check=('Check-'+$_);Status=$_;RunDateUtc=$now.AddHours(-$AgeHours).ToString('O')}
    }) | Export-Csv -LiteralPath $adPath -NoTypeInformation
}
function Read-Monitoring { Get-OptionalADHealthEvidence -Path $adPath -AsOfUtc $now }
function Get-Plan { @(Get-PreparedSourcePlan -DataRoot $raw -SourceContractPath $contract -MetadataOnly) }
$runParameters = @{DataRoot=$raw;OutputRoot=(Join-Path $raw 'DATA-POWERBI');WorkRoot=(Join-Path $fixture 'work');TenantKey='synthetic-test';AccountClassificationConfigPath=(Join-Path $fixture 'unused.json.txt');ValidateOnly=$true}

# Optional source selection and capture; all required and tenant guards stay strict.
$issues=[Collections.Generic.List[object]]::new()
$plan=@(Get-PreparedSourcePlan -DataRoot $raw -SourceContractPath $contract -OptionalSourceIssues $issues)
Check ($plan.Count -eq 1 -and $issues.Count -eq 1) 'Missing monitoring blocked selection or lost its audit'
Write-ADFixture
$plan=Get-Plan
$adEntry=@($plan | Where-Object IsOptional)[0]
Check ($plan.Count -eq 2 -and $adEntry.MaxAgeHours -eq 48) 'Fresh optional monitoring was not selected'
[IO.File]::SetLastWriteTimeUtc($adPath,$now.AddHours(-49))
Check ((Get-Plan).Count -eq 1) 'Stale publication was accepted as current monitoring'
Check (@(Get-PreparedSourcePlan -DataRoot $raw -SourceContractPath $contract -AgeOverrides @{'AD_HealthCheck.csv'=168}).Count -eq 1) 'Global age override weakened optional monitoring freshness'
[IO.File]::SetLastWriteTimeUtc($adPath,$now.AddMinutes(10))
Check ((Get-Plan).Count -eq 1) 'Future optional publication was accepted'
Write-ADFixture
Remove-Item -LiteralPath $adPath
$copyIssues=[Collections.Generic.List[object]]::new()
$snapshot=Join-Path $fixture 'capture'
$copy=& $module {param($entry,$root,$issues) Copy-PreparedStableSource -Entry $entry -SnapshotRoot $root -Attempts 1 -RetryDelayMs 0 -OptionalSourceIssues $issues} $adEntry $snapshot $copyIssues
Check ($null -eq $copy -and $copyIssues.Count -eq 1 -and -not (Test-Path (Join-Path $snapshot $adEntry.Relative))) 'Disappearing optional source blocked capture or retained partial bytes'
$result=Invoke-PreparedEvidencePipeline @runParameters
Check ($result.SourceFiles -eq 1 -and -not $result.Publication -and $result.OptionalSourceIssues.Count -eq 1) 'Missing optional source blocked ValidateOnly'
$capture=Get-Content -LiteralPath (Join-Path $result.DiagnosticPath 'capture.json.txt') -Raw | ConvertFrom-Json
Check ($capture.OptionalSourceIssues.Count -eq 1) 'Skipped monitoring was not retained in capture diagnostics'
Remove-Item -LiteralPath $requiredPath
Reject { Get-Plan } 'Required.csv'
[pscustomobject]@{TenantKey='synthetic-test';Value=1} | Export-Csv -LiteralPath $requiredPath -NoTypeInformation
[IO.File]::SetLastWriteTimeUtc($requiredPath,$now.AddHours(-169))
Reject { Get-Plan } 'publication-age'
[IO.File]::SetLastWriteTimeUtc($requiredPath,$now)
Write-ADFixture -Tenant 'other-test'
Reject { Invoke-PreparedEvidencePipeline @runParameters } 'TenantKey'

# Business acquisition freshness: modifying/copying a stale file never refreshes it.
Write-ADFixture
$monitoring=Read-Monitoring
Check ($monitoring.Available -and $monitoring.Rows.Count -eq 3 -and $monitoring.UnmeasuredChecks -eq 1) 'Measured coverage includes NotMeasured checks'
Write-ADFixture -AgeHours 48
Check ((Read-Monitoring).Available) 'Exactly 48-hour acquisition was incorrectly excluded'
Write-ADFixture -AgeHours 49
Check (-not (Read-Monitoring).Available) 'Recently copied stale acquisition was accepted'
Write-ADFixture -AgeHours -1
Check (-not (Read-Monitoring).Available) 'Future acquisition was accepted'
Write-ADFixture -Statuses @('NotMeasured')
Check (-not (Read-Monitoring).Available) 'Unmeasured checks became healthy evidence'
Write-ADFixture -Statuses @('Unexpected')
Check (-not (Read-Monitoring).Available) 'Unknown status became a healthy result'
Write-ADFixture
$rows=@(Import-Csv $adPath); $rows[0].RunDateUtc=''; $rows | Export-Csv $adPath -NoTypeInformation
Check (-not (Read-Monitoring).Available) 'Missing acquisition timestamp fell back to modification time'
Write-ADFixture
$rows=@(Import-Csv $adPath); $rows[1].RunDateUtc=$now.AddHours(-49).ToString('O'); $rows | Export-Csv $adPath -NoTypeInformation
Check (-not (Read-Monitoring).Available) 'Mixed observations used the newest timestamp'
'Forest,Domain,Category,Check,Status,RunDateUtc' | Set-Content $adPath
Check (-not (Read-Monitoring).Available) 'Header-only monitoring became observed evidence'
'Forest,Domain,Category,Check,Status,RunDateUtc','example.invalid,example.invalid,Replication,Check,OK' | Set-Content $adPath
Check (-not (Read-Monitoring).Available) 'Malformed monitoring became observed evidence'

# Execute both real readers; unchanged output schemas must pass prepared validation.
$device=Join-Path $fixture 'DeviceInventoryEvidence.csv'
$users=Join-Path $fixture 'UserInventoryEvidence.csv'
$security=Join-Path $fixture 'SecurityControlEvidence.csv'
$trust=Join-Path $fixture 'DataTrustEvidence.csv'
[pscustomobject]@{'Compliance State'='Compliant';'Secure Boot State'='Enabled'} | Export-Csv $device -NoTypeInformation
[pscustomobject]@{'MFA Registration State'='Registered'} | Export-Csv $users -NoTypeInformation
[pscustomobject]@{AzureADDeviceId='test-device';DeviceName='test-device';'BitLocker/Encryption'='Pass';CodeIntegrity='Pass'} | Export-Csv (Join-Path $last 'Intune_Devices_Compliance_Policies.csv') -NoTypeInformation
[pscustomobject]@{Status='OK'} | Export-Csv (Join-Path $last 'M365_Entra_AzureADConnect_SyncHealth.csv') -NoTypeInformation
foreach ($case in 'fresh','missing','stale','unmeasured') {
    if ($case -eq 'missing') { Remove-Item $adPath }
    elseif ($case -eq 'stale') { Write-ADFixture -AgeHours 49 }
    elseif ($case -eq 'unmeasured') { Write-ADFixture -Statuses @('NotMeasured') }
    else { Write-ADFixture }
    & (Join-Path $PSScriptRoot 'New-SecurityEvidence.ps1') -DataRoot $last -DeviceEvidencePath $device -UserEvidencePath $users -OutputPath $security -AsOfUtc $now | Out-Null
    $controls=@(Import-Csv $security)
    $control=@($controls | Where-Object 'Control Sort' -eq 5)[0]
    Check ($controls.Count -eq 11) "Control count changed: $case"
    if ($case -eq 'fresh') {
        Check ($control.'Covered Entities' -eq '3' -and $control.'Healthy Entities' -eq '1' -and $control.'Affected Entities' -eq '2') 'Measured AD counts changed'
        Check ([double]$control.'Health Rate' -lt 0.34) 'NotMeasured was counted as healthy'
    } else {
        Check ($control.'Evidence Status' -eq 'Not collected' -and $control.'Gap State' -eq 'Not assessed') "Unavailable monitoring became healthy: $case"
        foreach ($field in 'Covered Entities','Healthy Entities','Affected Entities','Health Rate','Snapshot Date') { Check ($control.$field -eq '') "Unavailable AD value fabricated: $field / $case" }
    }
    & (Join-Path $PSScriptRoot 'New-DataTrustEvidence.ps1') -DataRoot $raw -OutputPath $trust -AsOfUtc $now | Out-Null
    $adTrust=@(Import-Csv $trust | Where-Object 'File Name' -eq 'AD_HealthCheck.csv')[0]
    Check ($adTrust.'Decision Readiness' -eq $(if($case -eq 'fresh'){'Ready'}else{'Optional unavailable'})) "Data Trust still blocks optional monitoring: $case"
    foreach ($table in 'Security Control Evidence','Data Trust Evidence') {
        $validation=& (Get-Process -Id $PID).Path -NoProfile -File (Join-Path $PSScriptRoot 'Test-PreparedEvidence.ps1') -Root $fixture -Tables $table
        Check ($LASTEXITCODE -eq 0 -and ($validation | ConvertFrom-Json).Passed) "Prepared reader schema/type validation failed: $case / $table"
    }
}
Remove-Item (Join-Path $last 'M365_Entra_AzureADConnect_SyncHealth.csv')
Reject { & (Join-Path $PSScriptRoot 'New-SecurityEvidence.ps1') -DataRoot $last -DeviceEvidencePath $device -UserEvidencePath $users -OutputPath $security } 'Required private evidence'
$jobs=(Get-Content (Join-Path $PSScriptRoot '../../SmartM365/SmartInventory/Orchestrator/Orchestrator-Jobs.json.template') -Raw | ConvertFrom-Json).Jobs
Check ('AD-HealthCheck' -notin ($jobs | Where-Object Name -eq 'WorkplaceEvidence-Prepare').DependsOn) 'Workplace preparation still waits for AD health'
Check (($jobs | Where-Object Name -eq 'AD-HealthCheck').Enabled) 'AD monitoring was disabled'
Write-Host "PASS: $script:checks optional AD monitoring checks. Synthetic local evidence only; no collectors or live writes."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA9O/u/7WFw42h2
# h5MWIVOr4PrPPPpohTf/A8dK2CtiXqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIOp1wPe0Ii9YTKwb0HstOxtm5cnFLWf/MeV0/wQn0DUfMA0GCSqG
# SIb3DQEBAQUABIIBgDg6V+fZnbDvyQWNwQ91KIC0uAoCJzfqSO/2Ud5HP4qif70s
# oZ1TLeYUSb+AYWrcIzwfo746g15M268AGyGvTvSsoj0R204RNLtY+mVECqQj7BlE
# 8B2NWVV9ZC7GmEytibQoZBHerVMS+tEgXN4f11Ei9DRBt9v6b9X/v/gNWHqJhxoR
# C26UD/O/M+/kZlnCku/Cme85uBkbxwFoUbTYskODKaLjKxTdkNGsknxIazTNsnu2
# uqA2saYpa5J3cVs7w5kOlTtY8t0irAbUtF0Sr12EkOfwyWlrID5B4lMFM+jvX97L
# ES39xVeOtHue3ymE0PjrQaTZR9/oH6oDAnCl/IMmjKlxDtjx7BkXUl2oR8bTq0ow
# AFAvuSqb7P/T8mn91z6l+H2U3X24PBxOwGeQ5nFuCyE+tXKet8eEYKCXWR0YjN/R
# DVYIqGo8HYUIeShKiIoC7x94AxIZBoy3EU5Qf9brKVKWeW6gYnMkEO9W1bX4L4NN
# WkhUbqH8+iqGAl8vTaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDcxNjMz
# MzdaMC8GCSqGSIb3DQEJBDEiBCDzAleeVtnLtTi38DScibmpou6Kuu7F+hIPY7Ca
# e+SM+jANBgkqhkiG9w0BAQEFAASCAgBxCL6L/8oSHGzJeAKbzzTMA2eHsxnHw1Me
# E/7LpvD9sgtEOKogfF3bKOdutSbTQ6Pzj38/naN7/A9wfvy+53e7Hn1RapudbMUD
# xFoPl2ArWimD8tRHSZVu/A5FuOO23G+oDDtvnpXDNLNL/ieGOEHQ6xWRrDglKI5E
# 4K9Wh+6Jw6Tq089+PqSbFE/aTkB2+INyShbmhCTpuRT7NxDUFLsEeiew3jbdHgva
# tnOAqmMSEiGjpiP5/nDZDzVVFenmZeKIpuW6KkWGwi27EjbjKUtxuNxCmloh/XV/
# /Mi6BjPR0PVDCPs96In3jWJY84KicYlBMGdx5+1bChgspMrntctdahWDOY0fZGmo
# 53G+5w2tHGy2GeM/KM/urT9HXPA7fiYx3OFVFmoKZrU55X7AV6BCTCkmSjnFrtJb
# Ugk5sODSsK6OxGjr0ni5YvKirwkcAmaE7QsKwhZ0W3MEPT0KCQe43i6Fz/PkCHvN
# JBBhF9SvB5MKyM23MUvLyq6HWvJWM6gGgozeQrERlKCIvT99hEjO8QEea36bDZbH
# TIZ3DeRWu4XhD+jGRUROnPoTaXEaxnTyt90vh47VlAGxZRcIZSd9+3IP1WD4XrHi
# u9Sd1CVT1zdIjdWYuMtsgX3ph6S92iFvPSEi0uKTRtqFRnSEGmn6Nq9sTVKBxZFT
# 1YNYDW6fTw==
# SIG # End signature block
