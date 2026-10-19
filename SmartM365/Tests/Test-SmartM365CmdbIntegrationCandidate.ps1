#Requires -Version 7.0
<#
.SYNOPSIS
Offline contract tests for the retired CMDB integration candidate.
.DESCRIPTION
CmdbEvidence-Prepare is retired: the unified WorkplaceEvidence-Prepare job carries the
former CMDB producer dependencies. Reads repository contracts, merges jobs in memory and
calls the real pipeline selector and cutover planner. No tenant configuration, API,
submission, scheduling or file write.
.VERSION
1.1.0
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
$entryPath = Join-Path $inventoryRoot 'PreparedEvidence/SmartM365-CmdbEvidence-Orchestrator.ps1'
$hashes = @{}
foreach ($path in @($candidatePath,$jobsPath,$sourcePath,$intelligencePath,$entryPath)) { $hashes[$path] = (Get-FileHash -LiteralPath $path).Hash }
$candidate = Get-Content -LiteralPath $candidatePath -Raw | ConvertFrom-Json
$template = Get-Content -LiteralPath $jobsPath -Raw | ConvertFrom-Json
$producers = @((Get-Content -LiteralPath $sourcePath -Raw | ConvertFrom-Json).Producers)
$intelligence = Get-Content -LiteralPath $intelligencePath -Raw | ConvertFrom-Json
Import-Module (Join-Path $inventoryRoot 'Orchestrator/SmartM365.Orchestrator.Pipeline.psm1') -Force
Import-Module (Join-Path $inventoryRoot 'Orchestrator/SmartM365.WorkplaceCutover.psm1') -Force
function Assert-True { param([bool]$Condition, [string]$Label) if (-not $Condition) { throw $Label } }
function ConvertTo-CompactJson { param($Value) ConvertTo-Json -InputObject $Value -Depth 50 -Compress }
$testCount = 0
function Test-Case {
    param([string]$Name, [scriptblock]$Body)
    & $Body
    $script:testCount++
    Write-Output "PASS: $Name"
}
$retired = @($candidate.ProposedJobs | Where-Object Name -eq 'CmdbEvidence-Prepare')[0]
$unified = @($template.Jobs | Where-Object Name -eq 'WorkplaceEvidence-Prepare')
Test-Case 'Retired candidate stays an inactive review-only fixture' {
    Assert-True ($candidate.CandidateOnly -and -not $candidate.PSObject.Properties['Jobs']) 'Executable candidate manifest.'
    Assert-True (@($candidate.ProposedJobs).Count -eq 2) 'Unexpected integration expansion.'
    foreach ($job in $candidate.ProposedJobs) {
        Assert-True (-not $job.Enabled -and $job.AssignmentMode -eq 'Elected') 'Job activated or misleadingly Manual.'
        Assert-True (-not $job.PSObject.Properties['Schedule'] -and -not $job.PSObject.Properties['TimeoutMinutes']) 'Unqualified schedule/timeout.'
    }
    Assert-True (@($candidate.JobArgumentAdditions).Count -eq 1) 'Unrelated argument changes.'
    $addition = $candidate.JobArgumentAdditions[0]
    Assert-True ($addition.Name -eq 'Intune-DiscoveredApps-Inventory' -and $addition.Arguments -eq '-DeviceDetailMode All -FreshDeviceDetails') 'Fresh arguments missing.'
    foreach ($argument in @('-ResetResume','-MaxApps','-MaxItems','-DeviceDetailMode')) { Assert-True ($addition.RejectConflictingArguments -contains $argument) "Conflict not declared: $argument" }
    $tokens = $null; $errors = $null
    $entry = [Management.Automation.Language.Parser]::ParseFile($entryPath,[ref]$tokens,[ref]$errors)
    Assert-True ($errors.Count -eq 0) 'Retired CMDB entry point does not parse.'
    $first = $entry.EndBlock.Statements[0]
    Assert-True ($first -is [Management.Automation.Language.ThrowStatementAst] -and $first.Extent.Text -match 'CMDB activation is retired') 'Retired activation entry point can run.'
}
Test-Case 'Template keeps WorkplaceScope active and the retired CMDB job absent' {
    Assert-True (@($template.Jobs | Where-Object Name -eq 'CmdbEvidence-Prepare').Count -eq 0) 'Retired CMDB job remains in the active template.'
    Assert-True (@($template.Jobs | Where-Object { 'CmdbEvidence-Prepare' -in @($_.DependsOn) }).Count -eq 0) 'A job still depends on the retired CMDB job.'
    $scope = @($template.Jobs | Where-Object Name -eq 'M365-WorkplaceScope-Inventory')
    Assert-True ($scope.Count -eq 1 -and $scope[0].Enabled -and $scope[0].AssignmentMode -ne 'Manual' -and -not $scope[0].PSObject.Properties['RequiresExplicitActivation']) 'M365-WorkplaceScope-Inventory is not enabled in the template.'
    $proposedScope = @($candidate.ProposedJobs | Where-Object Name -eq 'M365-WorkplaceScope-Inventory')[0]
    Assert-True ($scope[0].ScriptPath -eq $proposedScope.ScriptPath) 'WorkplaceScope script path drifted.'
    Assert-True (Test-Path -LiteralPath (Join-Path $inventoryRoot $scope[0].ScriptPath) -PathType Leaf) 'WorkplaceScope script missing.'
    Assert-True ($unified.Count -eq 1 -and $unified[0].Enabled -and $unified[0].AssignmentMode -ne 'Manual' -and $unified[0].DependencyMode -eq 'FreshSuccess') 'Unified preparation is missing, disabled or ungated.'
}
Test-Case 'Unified preparation covers all 17 native producers' {
    $dependsOn = @($unified[0].DependsOn)
    Assert-True ($producers.Count -eq 17 -and @($retired.DependsOn).Count -eq 17) 'Producer registry changed; explicit review needed.'
    Assert-True (@($dependsOn | Sort-Object -Unique).Count -eq $dependsOn.Count) 'Duplicate dependency.'
    foreach ($name in $retired.DependsOn) { Assert-True ($dependsOn -contains $name) "Retired CMDB dependency not carried by the unified job: $name" }
    $dependencies = @(foreach ($name in $retired.DependsOn) {
        $matches = @($template.Jobs | Where-Object Name -eq $name)
        Assert-True ($matches.Count -eq 1 -and $matches[0].Enabled) "Missing, ambiguous or disabled CMDB producer job: $name"
        $matches[0]
    })
    # A producer script can back several jobs (EXO mailbox details and permissions).
    foreach ($producer in $producers) {
        Assert-True (@($dependencies | Where-Object { [IO.Path]::GetFileName($_.ScriptPath) -eq $producer.Script }).Count -ge 1) "Producer coverage: $($producer.Script)"
    }
    Assert-True ($dependsOn -contains 'EXO-Mailboxes-Inventory-Fast' -and $dependsOn -notcontains 'EXO-Mailboxes-Inventory') 'Acquisition must use daily mailbox details, not weekly stats.'
    $contract = Get-Content -LiteralPath (Join-Path $inventoryRoot 'PreparedEvidence/cmdb-prepared-contract.json.txt') -Raw | ConvertFrom-Json
    Assert-True ($contract.maxAgeHours -eq 48 -and $contract.maxCollectionSpanHours -eq 48) 'Core acquisition gates weakened.'
    Assert-True ($contract.freshnessGroups.Count -eq 1 -and $contract.freshnessGroups[0].maxAgeHours -eq 240 -and $contract.freshnessGroups[0].warningAgeHours -eq 168) 'Apps acquisition gates differ.'
}
Test-Case 'Real selector rejects the retired job and selects every producer for the unified job' {
    $rejected = $false
    try { Get-SmartM365OrchestratorPipelineSelection -JobsDocument $template -JobName CmdbEvidence-Prepare -IncludeDependencies | Out-Null }
    catch { $rejected = $_.Exception.Message -like '*absent from the jobs manifest*' }
    Assert-True $rejected 'Retired CMDB job can be requested.'
    $selection = Get-SmartM365OrchestratorPipelineSelection -JobsDocument $template -JobName WorkplaceEvidence-Prepare -IncludeDependencies
    foreach ($name in $retired.DependsOn) { Assert-True (@($selection.SelectedJobs | Where-Object Name -eq $name).Count -eq 1) "Not selected: $name" }
}
Test-Case 'A re-merged legacy CMDB job is removed by the cutover plan, enabled or disabled' {
    foreach ($enabled in @($false,$true)) {
        $merged = ConvertTo-CompactJson $template | ConvertFrom-Json -Depth 50
        $legacy = ConvertTo-CompactJson $retired | ConvertFrom-Json -Depth 50
        $legacy.Enabled = $enabled
        $merged.Jobs = @($merged.Jobs) + @($legacy)
        $plan = New-SmartM365WorkplaceCutoverPlan $merged
        Assert-True (@($plan.Jobs.Jobs | Where-Object Name -eq 'CmdbEvidence-Prepare').Count -eq 0) 'Re-merged CMDB job survived cutover.'
        Assert-True ((ConvertTo-CompactJson $plan.Jobs) -ceq (ConvertTo-CompactJson $template)) 'Cutover changed unrelated template jobs.'
    }
}
Test-Case 'Every Intelligence History/Trend retains its explicit stable key columns' {
    $dailyKey = 'TenantKey|OrganizationKey|EnvironmentKey|TenantId|Date Key|Source Name|Domain Key|Metric Name'
    $expected = @{
        'User Activity History Evidence'='Week Label'; 'Workforce Trend Evidence'='Week Label'
        'Windows Lifecycle Trend Evidence'='Snapshot Week'; 'Endpoint Experience Trend Evidence'='Snapshot Week'
        'Application Trend Evidence'='Snapshot Week'; 'Content Storage Trend Evidence'='Snapshot Week'
        'Collaboration Trend Evidence'='Snapshot Week|Service'; 'Executive KPI Trends'='Date Key|Metric Name'
        'Daily Directory Trends'=$dailyKey; 'Daily Mailbox Trends'=$dailyKey
    }
    $tables = @($intelligence.tables | Where-Object table -match 'History|Trend')
    Assert-True ($tables.Count -eq $expected.Count) 'History/Trend set changed; explicit review needed.'
    foreach ($table in $tables) {
        Assert-True ($table.PSObject.Properties['historyKey'] -and @($table.historyKey).Count -gt 0) "Missing historyKey: $($table.table)"
        Assert-True ((@($table.historyKey) -join '|') -eq $expected[$table.table]) "Stable key changed: $($table.table)"
        foreach ($key in $table.historyKey) { Assert-True (@($table.columns.name) -contains $key) "Key column missing: $($table.table)/$key" }
    }
}
Test-Case 'Inspected files unchanged' {
    foreach ($path in $hashes.Keys) { Assert-True ((Get-FileHash -LiteralPath $path).Hash -eq $hashes[$path]) "File mutated: $path" }
    Assert-True (@($candidate.ProposedJobs | Where-Object Enabled).Count -eq 0) 'Real candidate enabled.'
}
Write-Output "CMDB post-retirement contract tests: $testCount passed. No requests or external actions."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDlsFnVFuNKgh5r
# dU+dO2UdtjLGXdWUbDG6dQcLgUxFAqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEINkuhSmDYY77l1tEzsIEW0IH6OzBgO8BghTZ7FmveTq3MA0GCSqG
# SIb3DQEBAQUABIIBgBAvZCssuRnca2gGpKC8q0qhRUAEf/WH3iDJNabg7TJZwDGq
# nGn2OhHijeq4xiFudROxSv52GHRtvAZzvG8riY5LGCXTQ8b7jZwmmPmpRfuz7LCI
# +ZsJmmymJtsVtGTxx0Xj122eYvF+WJsgxYFNlxeBz2ZvH9+t7Qwf4sNkc3VrsVYM
# t6bSGsN0uDpvFoGCRAthzwUPH9GZw/ALz9q/yFSxj79unyEfCnRvc0rJaSBbvIuz
# Bkp1bNfXIgux03VL7gmHZ0zXHSKvAX9CUkFO8zNeArgXYhlIkBFXxfquvbu9hKI+
# R2XWdIQh+EeGHcPbXuG/Un0SBnk5ZnV3DIOzmp6e+30YBM9pW1MQh9DRuVyKJaXf
# GlG5Ip2oie7Paj5n/sQ/otD57JbN90exG3loZ1r9iDrFXgeNcw7UOKSNkEDi8yH8
# bg7LVjrGpNxn9JXWgPDNMIgwuEwRLR8Jtx7zAVQp3Zqqx/39zOk7rdnUp/wLlTCB
# 83W2AEfzPu2G8Eyr96GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTkxNTI2
# MDJaMC8GCSqGSIb3DQEJBDEiBCAa3G29B5C41nCYyvqpg0LAk9segk52RmXm63BA
# vRJ+hTANBgkqhkiG9w0BAQEFAASCAgBIQDT12fZAI7szCJ9EVkLaGSlf0x+GHrrP
# 9zZVsgMfiROH3Ar+PSVAQdSCHPjlj7+s0UM++WIH3VBQh360qjx0m1u8TMBRMD7E
# Hu6gHb+irNpd8q8j0v4mO+S32AKEVH4v1rOmzi6AAJpztJOroAKilnuUOyG/qrwo
# P4UtvgUIPHQfEXngq13b2IOb/SuuKC/xad9g4dLwYasjYTfGBUPeFAGHUA8ne/OL
# DEIFKqwSjX8SUyn0ThKl5lPTfKiFqiqg1+X+EB0KDkcdmscLMUdzkjo7jyLoS+89
# N2UvcYzSnmKXpo0xOG8GA7DDJAa7/n285sAauLoTbvVhLWiMmflEJQW+J8+VjMR6
# l3NQAke13R7VzU/Wb9cT7ycV39O+X2wo5eZuEHCqgzTCsjSE47TLqIQHwf2oFE+o
# c0KWl6bSfKIMuKEHmbpSrK3Fh3AwJcYqE54VvD8JPVwxz8813txFjZyqzvLFKC8J
# qaIDhdvNeDOUpGTGBMglPvdBvBiHk6br+iBB2074QjOyUfhhpU9XyxbAxDu3dBqW
# TpyYcd+cYePRTvlThyoYg85nS2yEKun3/PCcWosBK23uNSUmZGvGG1/ZLqVDYRb1
# Mgp01vdXBrS9ugHbVcBloi1NQowPtun+oRWrBUcjEnVkwosAkRgWm51uHR74N1SQ
# 78vvlOUs4w==
# SIG # End signature block
