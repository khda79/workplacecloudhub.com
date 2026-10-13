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
        Assert-True ($optIn.Count -eq 1) 'Template integration job missing.'
        Assert-True ($optIn[0].Enabled -and -not $optIn[0].PSObject.Properties['RequiresExplicitActivation']) "$($job.Name) is not enabled in the template."
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
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCsxmQU9XdJqtZv
# FA5h9whFnUUysZ+wST83e4Wlstdl56CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCiA7LmWCLL3AkErL8TxVFb
# U6/G1ctOQykvi9efwE2gjzANBgkqhkiG9w0BAQEFAASCAYBcF91VbFYzzbxpRDVi
# hvENUspev3GNPXf00+DkeaPGDqfRJDQV0cGJKO8SP+BcrTSqw/9Cqvx4cKGmRwWo
# J3gUFmF1W7mCha1HT6jW2pM1uloAQQ7qvxgPSGyOh5MleXLulspxlHqEaGM7Rt4m
# W1jdxc9nqfQz5V70VQbhl+okO+i+etVJyOu3AQ8i/MCky3eYHZOF5wb7YLE8RM9n
# a3f4W9s2uHvpPiSFoRZGFmox+25/dw13FrBBDZ1JO2wUZxc2zbWvwwil6pZOY84h
# k1yeFmbSx1XWDA6CnTvvd+uMGUWivCc6hmoOfPt16wWGxPUCM/hIDqdpUuOrrMwF
# llUSBBHi1q6mzxX48TOQCqJ438WNJ9gyXpeiEvkwthLCFkowpoY0KfL30LCFMUN9
# 8lO10MkQ7+F965USsQpjVWjamRgl6bJrQ5p7AoniEipnWt5MVf2WOj+H9Kz43iUB
# 63FMzrG3LVe2XuZhd47HSDoxA+vwrwSROb/rrpNZt1nzWmE=
# SIG # End signature block
