<#
.SYNOPSIS
Offline checks for a completed producer receipt used by FreshSuccess dependencies.
#>
#requires -Version 7.0
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$base = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
Import-Module (Join-Path $base 'SmartM365.Orchestrator.Distributed.psm1') -Force
Import-Module (Join-Path $base 'SmartM365.Orchestrator.Insights.psm1') -Force

$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $base 'SmartM365-Inventory-Orchestrator.ps1'), [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw 'Orchestrator source does not parse.' }
$definition = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Get-OrchestratorDependencyFreshStatus' }, $true).Extent.Text
$gate = New-Module -ScriptBlock ([scriptblock]::Create($definition))
$root = Join-Path $env:TEMP ('fresh-receipt-' + [guid]::NewGuid().ToString('N'))
$shared = Join-Path $root 'DATA-ALL/Orchestrator'
$latest = Join-Path $root 'DATA-LAST'
$claims = Join-Path $shared 'Election/Claims'
$scriptName = 'SmartM365-Synthetic-Inventory.ps1'
$receiptPath = Join-Path $latest 'SmartInventory_SmartM365-Synthetic-Inventory.current.json.txt'
$csvPath = Join-Path $latest 'Synthetic.csv'
$now = [datetime]'2026-10-08T11:23:00'
$job = [pscustomobject]@{ Name = 'Synthetic-Inventory'; ScriptPath = "Inventory/$scriptName"; Enabled = $true; AssignmentMode = 'Elected'; TimeoutMinutes = 60; Schedule = [pscustomobject]@{ Type = 'Weekly'; DaysOfWeek = @('Saturday'); Times = @('20:00') } }
$consumer = [pscustomobject]@{ Name = 'Consumer'; DependsOn = @($job.Name); DependencyMode = 'FreshSuccess'; DependencyMaxAgeHours = 48 }
$jobs = [pscustomobject]@{ Jobs = @($job, $consumer) }
$oldTenant = $global:SmartM365TenantKey
$script:passed = 0
function Check([bool]$Value, [string]$Message) { if (-not $Value) { throw $Message }; $script:passed++ }
function Write-Receipt($Document) { [IO.File]::WriteAllText($receiptPath, ($Document | ConvertTo-Json -Depth 6), [Text.UTF8Encoding]::new($false)) }
function GateStatus {
    & $gate { param($dependency, $time) Get-OrchestratorDependencyFreshStatus -DependencyJob $dependency -MaxAgeHours 170 -Now $time } $job $now
}
try {
    New-Item -ItemType Directory $claims, $latest -Force | Out-Null
    [IO.File]::WriteAllText($csvPath, 'TenantKey,Value' + [Environment]::NewLine + 'tenant-test,1')
    $global:SmartM365TenantKey = 'tenant-test'
    & $gate { param($path, $claimRoot) $script:Settings = [pscustomobject]@{ SharedDataFolderPath = $path; ElectionClaimsPath = $claimRoot }; $script:DependencyFreshCache = @{} } $shared $claims
    $runId = [guid]::NewGuid().ToString('N')
    $document = [ordered]@{
        Owner = 'SmartInventory-CmdbSourceReceipt'; Status = 'Completed'; TenantKey = 'tenant-test'; Producer = $scriptName
        RunId = $runId; StartedAtUtc = '2026-10-05T07:29:37Z'; CompletedAtUtc = '2026-10-06T06:53:17Z'; Errors = 0; IsPartialInventory = $false
        Files = @([ordered]@{ File = 'Synthetic.csv'; RunId = $runId; Producer = $scriptName; Status = 'Success'; Errors = 0; IsPartialInventory = $false; SHA256 = ('A' * 64) })
    }
    Write-Receipt $document
    Check ((GateStatus) -eq 'Ready') 'A current completed producer receipt did not release the runtime gate.'
    $readiness = @(Get-SmartM365OrchestratorDependencyReadiness -SharedDataFolderPath $shared -JobsDocument $jobs -JobName Consumer -Now $now)
    Check ($readiness.Count -eq 1 -and $readiness[0].State -eq 'Ready' -and $readiness[0].Detail -eq 'Qualified producer receipt') 'GUI readiness disagrees with the runtime gate.'
    $variant = [pscustomobject]@{ Name = 'Synthetic-Inventory-Variant'; ScriptPath = $job.ScriptPath }
    Check ($null -eq (Get-SmartM365OrchestratorFreshProducerReceipt -SharedDataFolderPath $shared -Job $variant -TenantKey 'tenant-test' -Now $now -MaxAgeHours 170)) 'A receipt for one script was accepted for another job variant.'

    $cases = @(
        @{ Name = 'other tenant'; Change = { $document.TenantKey = 'other-tenant' } },
        @{ Name = 'failed run'; Change = { $document.Status = 'Failed' } },
        @{ Name = 'partial result'; Change = { $document.IsPartialInventory = $true } },
        @{ Name = 'wrong producer'; Change = { $document.Producer = 'SmartM365-Other.ps1' } },
        @{ Name = 'stale completion'; Change = { $document.CompletedAtUtc = '2026-09-29T01:45:04Z' } },
        @{ Name = 'future completion'; Change = { $document.CompletedAtUtc = '2026-10-09T06:53:17Z' } },
        @{ Name = 'missing CSV'; Change = { $document.Files[0].File = 'Absent.csv' } }
    )
    foreach ($case in $cases) {
        $original = $document | ConvertTo-Json -Depth 6
        & $case.Change
        Write-Receipt $document
        & $gate { $script:DependencyFreshCache = @{} }
        Check ((GateStatus) -eq 'Waiting') "Rejected receipt accepted: $($case.Name)."
        $document = $original | ConvertFrom-Json -AsHashtable
    }
    # Shared contract 1.2: actual CSV hashes and the current publication fence are mandatory.
    $document.Owner = 'SmartInventory-SourceReceipt'; $document.ContractVersion = '1.2'
    $document.OrganizationKey = 'synthetic'; $document.EnvironmentKey = 'test'; $document.TenantId = 'synthetic-id'
    $document.Files[0].SHA256 = (Get-FileHash -LiteralPath $csvPath -Algorithm SHA256).Hash
    $sharedBaseline = $document | ConvertTo-Json -Depth 8
    foreach ($case in @(
        @{ Name='shared completed'; Ready=$true; Change={} },
        @{ Name='wrong contract'; Ready=$false; Change={$document.ContractVersion='9.0'} },
        @{ Name='unknown owner'; Ready=$false; Change={$document.Owner='Other'} },
        @{ Name='shared unversioned'; Ready=$false; Change={$document.Remove('ContractVersion')} },
        @{ Name='wrong hash'; Ready=$false; Change={$document.Files[0].SHA256=('B'*64)} },
        @{ Name='duplicate file'; Ready=$false; Change={$document.Files+=@($document.Files[0])} },
        @{ Name='unsafe path'; Ready=$false; Change={$document.Files[0].File='../Synthetic.csv'} },
        @{ Name='shared partial'; Ready=$false; Change={$document.IsPartialInventory=$true} },
        @{ Name='shared foreign tenant'; Ready=$false; Change={$document.TenantKey='other'} },
        @{ Name='shared failed'; Ready=$false; Change={$document.Status='Failed'} },
        @{ Name='shared stale'; Ready=$false; Change={$document.CompletedAtUtc='2026-09-29T01:45:04Z'} },
        @{ Name='missing protocol run'; Ready=$false; Change={$document.PublicationProtocol=1} }
    )) {
        $document = $sharedBaseline | ConvertFrom-Json -AsHashtable
        & $case.Change | Out-Null; Write-Receipt $document
        & $gate { $script:DependencyFreshCache = @{} }
        Check (((GateStatus) -eq 'Ready') -eq $case.Ready) "Shared receipt gate mismatch: $($case.Name)."
    }
    $document = $sharedBaseline | ConvertFrom-Json -AsHashtable
    $document.PublicationProtocol=1
    $runPath = Join-Path $latest 'SmartInventory_SmartM365-Synthetic-Inventory.run.json.txt'
    $run = @{Owner='SmartInventory-SourceRun';ContractVersion='1.0';RunId='new-attempt';Status='Collecting';PublicationStarted=$false;UnqualifiedPublication=$false}
    foreach ($field in @('TenantKey','OrganizationKey','EnvironmentKey','TenantId','Producer')) { $run[$field]=$document[$field] }
    foreach ($case in @(
        @{Status='Collecting';Fence=$false;Ready=$true},
        @{Status='Failed';Fence=$false;Ready=$true},
        @{Status='Publishing';Fence=$false;Ready=$false},
        @{Status='Failed';Fence=$true;Ready=$false},
        @{Status='Collecting';Fence=$true;Ready=$false},
        @{Status='Completed';Fence=$false;Ready=$true}
    )) {
        $run.Status=$case.Status; $run.UnqualifiedPublication=$case.Fence; $run.PublicationStarted=$case.Fence
        [IO.File]::WriteAllText($runPath, ($run | ConvertTo-Json -Depth 6))
        Write-Receipt $document; & $gate { $script:DependencyFreshCache = @{} }
        Check (((GateStatus) -eq 'Ready') -eq $case.Ready) "Publication fence mismatch: $($case.Status)/$($case.Fence)."
    }
    $run.TenantKey='foreign'
    [IO.File]::WriteAllText($runPath, ($run | ConvertTo-Json -Depth 6))
    & $gate { $script:DependencyFreshCache = @{} }
    Check ((GateStatus) -eq 'Waiting') 'Foreign publication fence was accepted.'
    "Fresh receipt offline checks passed: $script:passed"
}
finally {
    $global:SmartM365TenantKey = $oldTenant
    if ([IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($root)).TrimEnd('\') -eq [IO.Path]::GetFullPath($env:TEMP).TrimEnd('\') -and
        [IO.Path]::GetFileName($root) -like 'fresh-receipt-*') {
        Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD8Mxz0XV/v9IVB
# 99VlCaP+O/0vJ8x8ajZ3XR125m8DbaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEINsJ+kGKiyOkTmjOG7bbH4xXPFTVO97GIY4gQHbfP5giMA0GCSqG
# SIb3DQEBAQUABIIBgHNYlLibwXhJvu9V2pcH4CRliIEkHG0mdv/WLbfunUDsKjB6
# BVgY+6O1+b4z6I8h4/0lBEsX9Y0ipqLAhPvcU3/0WQi6rXUbggpH/Wjj2BPC2a37
# nWGgjvY7iwh2OHXp2aWWKCUvP031Gi8vV1rBnUJZMSKPpSG6J3ttntmX2LoyzGfw
# 2Nio2w46ZnHi8exZLCRSUy5OFT21G3bJYq1xFEMrbuoXhI4c9pS3Bh4maeVVCj/u
# o94Q8P7+ZYz73FCHV8ZLN6+lG8wUSw0cWyQJPko9qe5T0rUaOTNKjMJCiYHHc3cA
# vzmS7Dp+XAOBTP4M9rCy2v0BXtJ8/dEZVeCPhpFH1fRvmUXnMpUSGqM0GVRWa8B6
# /nabW6IH+mqc8PjTaYXXyHa21obN0qUzFL+MkFkVM2rjJVTXMtTT0kEv4IW69075
# PER6YjjdoDwjmRCRa+SdlZgPowbwgdKQXElWfbvdYnXiNpahbRMMQm89kElHJAwi
# 2S5KDQ3pJUDPFmp306GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTQyMjE1
# NTVaMC8GCSqGSIb3DQEJBDEiBCBdBG3B7Gd25ILGRRAnGtg9xKuyv6ZgCV6Z2LaB
# W1wE+DANBgkqhkiG9w0BAQEFAASCAgCULH6TaIigZlyJSggJkYNom0u6PzI5f4RD
# +ldDWVPZzDFUciKsA91HlZfI1Y9b9MbMQ5eN5hx1ajDmaXwKJvpAquzClDoD2SUR
# DyLuORFq/wtv7qLV+YrP1eShfWgV2E9xH2YN9Gh7GTY/YpxH1NQeNAPoYDgrB0sk
# +wPICYh3x/loltk8DdmlUm0ApbGOeeKgfXSn4/zax3mAfOSYmgTQV0ouhUHSzTha
# 2zQL7BXWWArxgWDE7mwd74tSt6HPAuCWpO3KMCwsOudmrL0bLvktEajYrWqZs/Sz
# tsnJOrFWN4nmP3xImETyur8XQdt1eSGSvOKs+7MwMrH4w6GINwCHm9gCujgUch2F
# 4g6DE3UMgW67+eNcd6sXpbREd5WWchn8tbyXyDa/X2A10FbwkdKWN9gQcDI+pLKT
# gA0UKhmYrhqVfhnneH0aTVdnPBqLTkr5pdJSG50hKUtdulT55GeP7BvP55ORRGxX
# M8KzBQfqe+7sV1vv6sZehNrZpL4iAise7TTO66tASzZ5rgUtdkcD93Iuhn8gp2HY
# qZW1x1Q8qPhGMoRO3N3qkJaud/tj34twKG/LrQoJns8Y2fAtcxYGX4Khlgh7Jqcx
# vRzcJuMkYuKEXg7ZxcXoS3fTfg1xAu0sn++nwVzAkh/zRahbpX9JYw31WkZ5LHUH
# hLcclLC9ig==
# SIG # End signature block
