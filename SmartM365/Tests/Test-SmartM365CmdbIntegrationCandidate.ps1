#Requires -Version 7.0
<#
.SYNOPSIS
Offline contract and dependency tests for the inactive CMDB integration candidate.
.DESCRIPTION
Reads repository contracts, merges jobs in memory and calls the real pipeline
selector. No tenant configuration, API, submission, scheduling or file write.
.VERSION
1.0.2
#>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$smartRoot = Split-Path $PSScriptRoot -Parent
$repoRoot = Split-Path $smartRoot -Parent
$inventoryRoot = Join-Path $smartRoot 'SmartInventory'
$candidatePath = Join-Path $inventoryRoot 'PreparedEvidence/cmdb-orchestrator-integration.json.txt'
$jobsPath = Join-Path $inventoryRoot 'Orchestrator/Orchestrator-Jobs.json.template'
$sourcePath = Join-Path $smartRoot 'Modules/SmartM365.Core/SmartM365-CmdbSources.json.txt'
$intelligencePath = Join-Path $repoRoot 'SmartWorkplaceIntelligence/config/prepared-evidence-contract.json.txt'
$hashes = @{}
foreach ($path in @($candidatePath,$jobsPath,$sourcePath,$intelligencePath)) { $hashes[$path] = (Get-FileHash -LiteralPath $path).Hash }
$candidate = Get-Content -LiteralPath $candidatePath -Raw | ConvertFrom-Json
$originalJobs = Get-Content -LiteralPath $jobsPath -Raw | ConvertFrom-Json
$producers = @((Get-Content -LiteralPath $sourcePath -Raw | ConvertFrom-Json).Producers)
$intelligence = Get-Content -LiteralPath $intelligencePath -Raw | ConvertFrom-Json
Import-Module (Join-Path $inventoryRoot 'Orchestrator/SmartM365.Orchestrator.Pipeline.psm1') -Force
function Assert-True { param([bool]$Condition, [string]$Label) if (-not $Condition) { throw $Label } }
$testCount = 0
function Test-Case {
    param([string]$Name, [scriptblock]$Body)
    & $Body
    $script:testCount++
    Write-Output "PASS: $Name"
}
Test-Case 'Inactive specification cannot replace or activate the real manifest' {
    Assert-True ($candidate.CandidateOnly -and -not $candidate.PSObject.Properties['Jobs']) 'Executable candidate manifest.'
    Assert-True (@($candidate.ProposedJobs).Count -eq 2) 'Unexpected integration expansion.'
    foreach ($job in $candidate.ProposedJobs) {
        Assert-True (-not $job.Enabled -and $job.AssignmentMode -eq 'Elected') 'Job activated or misleadingly Manual.'
        Assert-True (-not $job.PSObject.Properties['Schedule'] -and -not $job.PSObject.Properties['TimeoutMinutes']) 'Unqualified schedule/timeout.'
        $optIn = @($originalJobs.Jobs | Where-Object Name -eq $job.Name)
        Assert-True ($optIn.Count -eq 1 -and -not $optIn[0].Enabled -and $optIn[0].RequiresExplicitActivation) 'Template integration job can activate without explicit review.'
        Assert-True (Test-Path -LiteralPath (Join-Path $inventoryRoot $job.ScriptPath) -PathType Leaf) 'Candidate script missing.'
    }
}
$merged = $originalJobs | ConvertTo-Json -Depth 30 | ConvertFrom-Json
$merged.Jobs = @($merged.Jobs | Where-Object Name -notin @($candidate.ProposedJobs.Name))
$merged.Jobs = @($merged.Jobs) + @($candidate.ProposedJobs | ConvertTo-Json -Depth 15 | ConvertFrom-Json)
$prepare = @($merged.Jobs | Where-Object Name -eq 'CmdbEvidence-Prepare')[0]
Test-Case 'All 17 native producers have exactly one dependency' {
    Assert-True ($producers.Count -eq 17 -and @($prepare.DependsOn).Count -eq 17) 'Producer/dependency count differs.'
    Assert-True (@($prepare.DependsOn | Sort-Object -Unique).Count -eq 17) 'Duplicate dependency.'
    $dependencies = @(foreach ($name in $prepare.DependsOn) {
        $matches = @($merged.Jobs | Where-Object Name -eq $name)
        Assert-True ($matches.Count -eq 1) "Missing or ambiguous job: $name"
        $matches[0]
    })
    foreach ($producer in $producers) {
        Assert-True (@($dependencies | Where-Object { [IO.Path]::GetFileName($_.ScriptPath) -eq $producer.Script }).Count -eq 1) "Producer coverage: $($producer.Script)"
    }
    Assert-True ($prepare.DependsOn -contains 'EXO-Mailboxes-Inventory-Fast' -and $prepare.DependsOn -notcontains 'EXO-Mailboxes-Inventory' -and $prepare.DependsOn -notcontains 'EXO-Mailboxes-Permissions') 'CMDB acquisition must use daily mailbox details, not weekly stats or permissions-only.'
    Assert-True ($prepare.DependencyMode -eq 'FreshSuccess' -and $prepare.DependencyMaxAgeHours -eq 240) 'Weekly Apps scheduler gate differs.'
    $contract=Get-Content -LiteralPath (Join-Path $inventoryRoot 'PreparedEvidence/cmdb-prepared-contract.json.txt') -Raw | ConvertFrom-Json
    Assert-True ($contract.maxAgeHours -eq 48 -and $contract.maxCollectionSpanHours -eq 48) 'Core acquisition gates weakened.'
    Assert-True ($contract.freshnessGroups.Count -eq 1 -and $contract.freshnessGroups[0].maxAgeHours -eq 240 -and $contract.freshnessGroups[0].warningAgeHours -eq 168) 'Apps acquisition gates differ.'
}
Test-Case 'Real selector rejects disabled CMDB candidate' {
    $rejected = $false
    try { Get-SmartM365OrchestratorPipelineSelection -JobsDocument $merged -JobName CmdbEvidence-Prepare -IncludeDependencies | Out-Null }
    catch { $rejected = $_.Exception.Message -like '*disabled or manual*' }
    Assert-True $rejected 'Disabled candidate can be requested.'
}
Test-Case 'Synthetic enabled copy selects preparation and all 17 producers' {
    foreach ($name in @('M365-WorkplaceScope-Inventory','CmdbEvidence-Prepare')) { @($merged.Jobs | Where-Object Name -eq $name)[0].Enabled = $true }
    $selection = Get-SmartM365OrchestratorPipelineSelection -JobsDocument $merged -JobName CmdbEvidence-Prepare -IncludeDependencies
    Assert-True (@($selection.SelectedJobs).Count -eq 18 -and @($selection.IgnoredDependencies).Count -eq 0) 'Required producer omitted.'
    foreach ($name in $prepare.DependsOn) { Assert-True (@($selection.SelectedJobs | Where-Object Name -eq $name).Count -eq 1) "Not selected: $name" }
}
Test-Case 'Existing jobs and Intelligence dependencies unchanged; fresh arguments explicit' {
    foreach ($original in @($originalJobs.Jobs | Where-Object Name -notin @($candidate.ProposedJobs.Name))) {
        $copy = @($merged.Jobs | Where-Object Name -eq $original.Name)[0]
        Assert-True (($original | ConvertTo-Json -Depth 15 -Compress) -eq ($copy | ConvertTo-Json -Depth 15 -Compress)) "Original job changed: $($original.Name)"
    }
    Assert-True (@($candidate.JobArgumentAdditions).Count -eq 1) 'Unrelated argument changes.'
    $addition = $candidate.JobArgumentAdditions[0]
    Assert-True ($addition.Name -eq 'Intune-DiscoveredApps-Inventory' -and $addition.Arguments -eq '-DeviceDetailMode All -FreshDeviceDetails') 'Fresh arguments missing.'
    foreach ($argument in @('-ResetResume','-MaxApps','-MaxItems','-DeviceDetailMode')) { Assert-True ($addition.RejectConflictingArguments -contains $argument) "Conflict not declared: $argument" }
}
Test-Case 'Every Intelligence History/Trend retains its explicit stable key columns' {
    $expected = @{
        'User Activity History Evidence'='Week Label'; 'Workforce Trend Evidence'='Week Label'
        'Windows Lifecycle Trend Evidence'='Snapshot Week'; 'Endpoint Experience Trend Evidence'='Snapshot Week'
        'Application Trend Evidence'='Snapshot Week'; 'Content Storage Trend Evidence'='Snapshot Week'
        'Collaboration Trend Evidence'='Snapshot Week|Service'; 'Executive KPI Trends'='Date Key|Metric Name'
    }
    $tables = @($intelligence.tables | Where-Object table -match 'History|Trend')
    Assert-True ($tables.Count -eq $expected.Count) 'History/Trend set changed; explicit review needed.'
    foreach ($table in $tables) {
        Assert-True ($table.PSObject.Properties['historyKey'] -and @($table.historyKey).Count -gt 0) "Missing historyKey: $($table.table)"
        Assert-True ((@($table.historyKey) -join '|') -eq $expected[$table.table]) "Stable key changed: $($table.table)"
        foreach ($key in $table.historyKey) { Assert-True (@($table.columns.name) -contains $key) "Key column missing: $($table.table)/$key" }
    }
}
Test-Case 'Inspected files unchanged; synthetic enabling did not escape its copy' {
    foreach ($path in $hashes.Keys) { Assert-True ((Get-FileHash -LiteralPath $path).Hash -eq $hashes[$path]) "File mutated: $path" }
    Assert-True (@($candidate.ProposedJobs | Where-Object Enabled).Count -eq 0) 'Real candidate enabled.'
}
Write-Output "CMDB integration candidate tests: $testCount passed. No requests or external actions."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDUKoOBRvvnJ68l
# VFSde0RbfH18uhPZPep6ehDjq+WnxaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIG4QIvV9nUlMGYwIYlM4VhzL5oyb2BrUPKKK6Ol4LVHgMA0GCSqG
# SIb3DQEBAQUABIIBgCvswpO1Q9DL2uLhYVjKqef5Rh5ZPTgQtX3cRz5AkCVTZ2BI
# cOYB1v5uaCVCHgerefOKvmsQvjLkGuorav4EkUa3mv1/8ObMAryAnXjVA9OmfXfM
# 5kOO9cvnDO/bntssSsqojGVo+rdF7Z9EOa4cV1Ys6kmT52M7KWjXS64++RjPDpsO
# JeBz8wiJiKf3iAqDh1H4QkOr8efD2rsyacTP36f+tyrmhOGp1nIMyfH2z4pPqsZ7
# m/bnoB6GPQ6O25IGyyNduGa5U+Mi8YPa3wYgIB+hqIAt7cyR8VGsYkVHN13TtXkD
# alV322J+0zLYvdSKXehN+TJS3N/cLIkl4LF/6iU8Qe1a0CGE6j9GjYhd7kJ5pRvL
# 5vhdn0XvtwUF20BqkIQ8txtfsiI3JmrTi1Ww/wocQ775jKEhm+gwentWkAG9tPWp
# dSc/0Zs46F8Ui+sOuwyNPJaOyfLMNID+589Fu6uba8PauiX8YvkI6CqD+VsnaMUc
# XuPny1kNvizPTwjSvKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDcxNDQ1
# NDlaMC8GCSqGSIb3DQEJBDEiBCCaCEXpiQ5k67OG08j2UmWaHQqpq4oMwwRYAHdV
# RNLt3zANBgkqhkiG9w0BAQEFAASCAgCDAinNTnEw0VDC0SmPT82X8QSHnoZJg8HH
# UFqmgxEmWtfoq2ZggRx+SE6CKfOjHo4H1aXVXctS5KWZx2GK/qBMOCUQByFRV8SP
# eqEvQ6GEfhMymuU8Z5OpNJrZttv37OuZYwoLuXp1ujsMF0s7PRmg0g0AAXADtOCf
# xIU4NVWVWrtlqI9/gzyqBJ8oG3dmUTDFQyM8JU6021KjdF8xEDynLtLCJ4uiyuC1
# sUJg+tuORhDCHsoaNImEYChGszoL8RVyfki1hi4j2FNKZFKR/Ldy3t+UsUSFRTjT
# VtGJm5Kl6h+3V8eKmnHER3yxXHWmzVOLHP/YTkX8UTbb8nz3Jf0IQsuR3H2rKADA
# L2aBBVG4pvheheixTxmYOLdZ1S0PNFrYJKF+Jd1abnpzr6R5+JAzH9yx/dABEXGH
# 3jn8VMpWdWFscPJHyzRB768vn1aVRUxcOxMmfE2VNaAQFaSPlWJRBgewDXd/grKf
# BBOkiQWp5mFfWR8lmWVuanLEl+EdY/sADv5804FzOtrvdyxVTq3PitbI0vCWYi4W
# JB3M+zhap6AGijs0TiKDP9bNl1Lbkj418ZOG+3RSy6klfnLirYH5F2bhGEePY4Ly
# oX864JlgpxYglqlpQbXl5RoRH6a7jykSRwsJJ0d11xgvDT5Q3wQuX4C+Hp5Fh+vB
# FOWGVe1/YQ==
# SIG # End signature block
