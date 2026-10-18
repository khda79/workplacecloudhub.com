#requires -Version 5.1
<#
.SYNOPSIS
Validates synthetic application behavior against the current source version.
.VERSION
1.0.1
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Assert-True {
    param([bool]$Condition,[string]$Message)
    if (-not $Condition) { throw "ASSERT FAILED: $Message" }
}

function Assert-Equal {
    param($Actual,$Expected,[string]$Message)
    if ([string]$Actual -ne [string]$Expected) {
        throw "ASSERT FAILED: $Message. Expected='$Expected'; Actual='$Actual'."
    }
}

$toolkitRoot = Split-Path -Parent $PSScriptRoot
$launcherPath = Join-Path $toolkitRoot 'Scripts\SmartM365-Invoke-IntuneHybridJoinRepairWithPsExec.ps1'
$endpointPath = Join-Path $toolkitRoot 'Scripts\SmartM365-Invoke-IntuneHybridJoinRepair.ps1'

$tokens = $null
$parseErrors = $null
$launcherAst = [System.Management.Automation.Language.Parser]::ParseFile($launcherPath,[ref]$tokens,[ref]$parseErrors)
if ($parseErrors.Count -gt 0) { throw "Launcher parsing failed: $($parseErrors[0].Message)" }

foreach ($functionName in @('Get-LocalWorkerStartDiagnosticText','Start-LocalWorkerJobWithRetry','Get-ComputerListKey','Get-PostCycleCloudRefreshRows','Merge-ScopedInventoryMap','Test-AdInventoryRefreshDue','Get-AdaptiveCycleDelaySeconds')) {
    $functions = @($launcherAst.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $functionName
    },$true))
    Assert-Equal $functions.Count 1 "helper $functionName exists exactly once"
    . ([scriptblock]::Create($functions[0].Extent.Text))
}

$global:SyntheticWorkerStartAttempts = 0
$recoveredWorkerStart = Start-LocalWorkerJobWithRetry -ComputerName 'CH-RETRY-001' -JobName 'SyntheticRetry' -MaxAttempts 3 -RetryDelaySeconds 0 -StartOperation {
    $global:SyntheticWorkerStartAttempts++
    if ($global:SyntheticWorkerStartAttempts -lt 3) { throw 'Synthetic transient Start-Job failure.' }
    [pscustomobject]@{ Id = 42; Name = 'SyntheticRetry' }
}
Assert-True $recoveredWorkerStart.Succeeded 'local worker start recovers on the third attempt'
Assert-Equal $global:SyntheticWorkerStartAttempts 3 'local worker start retries exactly twice before recovery'
Assert-Equal $recoveredWorkerStart.Job.Id 42 'recovered local worker job is returned'

$global:SyntheticWorkerStartAttempts = 0
$failedWorkerStart = Start-LocalWorkerJobWithRetry -ComputerName 'CH-FAIL-001' -JobName 'SyntheticFailure' -MaxAttempts 3 -RetryDelaySeconds 0 -StartOperation {
    $global:SyntheticWorkerStartAttempts++
    throw 'Synthetic persistent Start-Job failure.'
}
Assert-True (-not $failedWorkerStart.Succeeded) 'persistent local worker start failure is returned without throwing'
Assert-Equal $global:SyntheticWorkerStartAttempts 3 'persistent local worker start failure stops after three attempts'
Assert-True ($failedWorkerStart.Detail -match 'Attempt=3/3') 'persistent failure records every attempt'
Assert-True ($failedWorkerStart.Detail -match 'ExpectedExecutable=.*pwsh\.exe|ExpectedExecutable=.*powershell\.exe') 'persistent failure records the expected PowerShell executable'
Assert-True ($failedWorkerStart.Detail -match 'FileExists=') 'persistent failure records File.Exists evidence'
Assert-True ($failedWorkerStart.Detail -match 'ProcessHandleCount=') 'persistent failure records process handle evidence'
Remove-Variable -Name SyntheticWorkerStartAttempts -Scope Global -ErrorAction SilentlyContinue

$cloudCandidates = @(Get-PostCycleCloudRefreshRows -Rows @(
    [pscustomobject]@{ Computer='FR-01'; Status='ADMIN_SHARE_UNREACHABLE' },
    [pscustomobject]@{ Computer='FR-02'; Status='SKIPPED_BY_STATUS_BACKOFF' },
    [pscustomobject]@{ Computer='FR-03'; Status='DRYRUN_READY' },
    [pscustomobject]@{ Computer='FR-04'; Status='WAITING_FOR_INTUNE_ENROLLMENT' }
))
Assert-Equal $cloudCandidates.Count 1 'only cloud-change-capable rows enter the post-cycle Graph scope'
Assert-Equal $cloudCandidates[0].Computer 'FR-04' 'the actionable device remains in scope'

$existingMap = @{
    'FR-SCOPE' = [pscustomobject]@{ Marker='stale' }
    'FR-KEEP' = [pscustomobject]@{ Marker='keep' }
    '__SMARTM365_INVENTORY_CHECKED__' = [pscustomobject]@{}
}
$refreshedMap = @{
    'FR-SCOPE' = [pscustomobject]@{ Marker='fresh' }
    '__SMARTM365_INVENTORY_CHECKED__' = [pscustomobject]@{}
}
$mergedMap = Merge-ScopedInventoryMap -ExistingMap $existingMap -RefreshedMap $refreshedMap -ScopedComputers @('fr-scope.contoso.test','FR-ABSENT')
Assert-Equal $mergedMap.Count 2 'scoped merge retains unrelated inventory and removes scoped absences'
Assert-Equal $mergedMap['FR-SCOPE'].Marker 'fresh' 'scoped refreshed device replaces stale data'
Assert-Equal $mergedMap['FR-KEEP'].Marker 'keep' 'unrelated cached device is preserved'
Assert-True (-not $mergedMap.ContainsKey('__SMARTM365_INVENTORY_CHECKED__')) 'inventory sentinel is not merged into the complete map'

$now = [datetime]::SpecifyKind([datetime]'2026-07-29T12:00:00',[DateTimeKind]::Utc)
Assert-True (Test-AdInventoryRefreshDue -LastRefreshUtc $null -FreshnessHours 12 -NowUtc $now) 'missing AD refresh timestamp is due'
Assert-True (-not (Test-AdInventoryRefreshDue -LastRefreshUtc $now.AddHours(-11) -FreshnessHours 12 -NowUtc $now)) 'fresh AD cache is reused'
Assert-True (Test-AdInventoryRefreshDue -LastRefreshUtc $now.AddHours(-12) -FreshnessHours 12 -NowUtc $now) 'AD cache refreshes when TTL expires'

$allBackoffRows = @(
    [pscustomobject]@{ Status='SKIPPED_BY_STATUS_BACKOFF'; EffectiveStatus='SKIPPED_BY_STATUS_BACKOFF'; BackoffUntilUtc=$now.AddMinutes(90).ToString('o') },
    [pscustomobject]@{ Status='SKIPPED_BY_TECH_RUN_GUARD_STARTED_NO_RESULT'; EffectiveStatus='SKIPPED_BY_TECH_RUN_GUARD_STARTED_NO_RESULT'; BackoffUntilUtc=$now.AddMinutes(120).ToString('o') }
)
$adaptiveDelay = Get-AdaptiveCycleDelaySeconds -Rows $allBackoffRows -MinimumDelaySeconds 60 -NowUtc $now
Assert-Equal $adaptiveDelay 5400 'all-backoff cycle waits for the earliest future expiry'
$mixedDelay = Get-AdaptiveCycleDelaySeconds -Rows @($allBackoffRows[0],[pscustomobject]@{ Status='ACTIONABLE'; EffectiveStatus='ACTIONABLE'; BackoffUntilUtc='' }) -MinimumDelaySeconds 60 -NowUtc $now
Assert-Equal $mixedDelay 60 'an actionable row preserves the configured minimum delay'
$invalidDelay = Get-AdaptiveCycleDelaySeconds -Rows @([pscustomobject]@{ Status='SKIPPED_BY_STATUS_BACKOFF'; EffectiveStatus='SKIPPED_BY_STATUS_BACKOFF'; BackoffUntilUtc='invalid' }) -MinimumDelaySeconds 60 -NowUtc $now
Assert-Equal $invalidDelay 60 'invalid backoff evidence safely falls back to the configured delay'

$launcherText = Get-Content -LiteralPath $launcherPath -Raw
foreach ($requiredText in @('$LauncherVersion = "2.10.78"','CloudRefreshScope.txt','Merge-ScopedInventoryMap','remoteStagingScript','Move-Item -LiteralPath $remoteStagingScript','LOCAL_WORKER_START_FAILED')) {
    Assert-True $launcherText.Contains($requiredText) "launcher source contains $requiredText"
}

$endpointTokens = $null
$endpointErrors = $null
$endpointAst = [System.Management.Automation.Language.Parser]::ParseFile($endpointPath,[ref]$endpointTokens,[ref]$endpointErrors)
if ($endpointErrors.Count -gt 0) { throw "Endpoint parsing failed: $($endpointErrors[0].Message)" }
$dsregFunctions = @($endpointAst.FindAll({
    param($node)
    $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Get-ParsedDsregStatusSnapshot'
},$true))
Assert-Equal $dsregFunctions.Count 1 'dsreg snapshot helper exists exactly once'
$dsregText = $dsregFunctions[0].Extent.Text
Assert-True (-not $dsregText.Contains('$env:TEMP')) 'dsreg snapshot no longer uses the shared Windows temp directory'
Assert-True $dsregText.Contains('$OutputDirPath') 'dsreg snapshot writes directly into the run evidence directory'
Assert-True (Get-Content -LiteralPath $endpointPath -Raw).Contains('$ScriptVersion = "2.10.40"') 'endpoint version is incremented'

Write-Output 'PASS: Intune Hybrid Join orchestrator optimization tests completed.'

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDcHveMoj50xfTf
# Kgtif7GuS//AZpyxbJkAHyuEm80hRqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIGRo7VHtKRKQRiCCf2O9KEaksj+xWaeMSDNs8VZsbDT+MA0GCSqG
# SIb3DQEBAQUABIIBgBf8Pxhgu29tW5weEDd0JuBssrAoqSWe2LufczfTSqkZXpn9
# ny41UrcWFf6bcaK8Bo9UUgggCV9rl8uu/h0NalZb9rtm+0zE7d9V3VFFQnPnO9HR
# FErgPlq329m4s82R255TSda3DfxQ7UEDKzA+o+v9BrBYRlRNBYLnBxmNkIBOS7wu
# DcQw9ucLYmqScp6qo+3ZbgudHXrYzI3vgDW5brCJpua5nBvzCi5Z2PSoTCppaniG
# wmyeLehWaenUS25Jy032vZIAj5kNHfgTG+6xputcfyOWAL1yGeZoWLFa+2D/9B52
# dfci/nRCxdXb8kTdgoNml00Eh0QmAk3ArHOH7F5jIVZBUOgPsOrkZiwlaZO45Joh
# GWzFSK7/mWFPrvLBxIJu/+RXUwDUyW/QRWCTl8K1UGYQEFtIP/XU3uH8O1nuAo4j
# Yr5QHN5dU/rXFPSHRhzrK6WYxCzbJ/Z0ADn15dHKG4bv248qAFVkY9VIEmaq8HG1
# lYWATcu2K0PvWMFHkqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTgwMzI1
# MDhaMC8GCSqGSIb3DQEJBDEiBCDFtxOZFp3qGEZzC0WB0uKlNrEHycdiOkj9HhK1
# s7LSNDANBgkqhkiG9w0BAQEFAASCAgCl68dneOId3O4fX5Q03YLSMH/UEG5kbBLt
# OIZ2Yx5ePNR77czxWR291nascfhNZLlDNLgGYLfO2yTNHu9DhUUb3aYiXurmmvbu
# o/xdb4vdt/SiGwox6K8BLeaH1T9QFfogXTb4ZA81sfn7JIpU/46EKKccTMlBQqzt
# RvPBuKMPbi/Wo5SWotQkn4skMfYO0Z7Y8acOZZm0AOG6376sLPqyDxCFWLg47gjw
# i2PaDotLhP5QfSxcYmc0qOt1pL+pb9zddQlueV0PYN9riXoEHzjpj7L+vfWWNM3n
# i7vv+hzB43nd+nRMQJVuzSXJBxjaeZDQJyBw/bjASoxgKr6GORFsuzI8Ke2l0VCh
# My4x+RP8gs5+ZK2OdCOvj+IfRIeGSgWLPwmUXDgwCP7z5FZynJCZwFx51T5sMx1K
# +S901N/PRw0BXlAnQnepH5d1cfyJbnbiPknD042trrtcVgOtMMka+StLHyq4uyBT
# ERJ3EjsxFTzY6jOIFNDPl23pvnxVPvA6oZEJm3BcuROXvONeKlFAUHgrAKUTrekp
# V0lDg7OFuUFkdEYrIOwXUXb2+TR+lnbBcnFav4fQ7E6XIK0qenwT0Cfs4L2Xo6tC
# r/eSjTBi0h+cbYDPrFfTpLsM119VdBy9wfUbVsxzF66bOGkPPdzY6qbZcT/FLhQo
# erkXGQiEhQ==
# SIG # End signature block
