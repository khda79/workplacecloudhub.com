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
    "Fresh receipt offline checks passed: $script:passed"
}
finally {
    $global:SmartM365TenantKey = $oldTenant
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDxvSb+UAOoij81
# UscvUzAG0mEuHSwlOpCKMCZuNHt25aCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCR5fUHq6lLYjPRIZ44V4MP
# iAyKnAzJtgVHvjhHJsgEmTANBgkqhkiG9w0BAQEFAASCAYBZLv8KEAjxxOFPZZzo
# lVPMeaNE8pt/aAy1A1KiH/lbGd+EMVMEr069JG4w7E0jZNeYfr4PDcS2aFwcR1ri
# o+DXc1bgHXqryCMT5sVwC32n7Wh6sTT3CmT4FdxrkGFooycCr2Vvw3LhUJUN33rG
# dnuHOlQrCnej9jhzGwnJxrRo7EqSngYk9xNmANEZbcNI8vFE192VKWvLU5shjI/f
# h39aLmpcXGMn6rvCj+odhH/4CuFrUuLGB1axIr+gf/HpiUUUk/ueLC3SKaSj4YYr
# Tw/87tBCT5prXPxJpEx4Ea7iEnGciTYwyDwaIBwJYp7sWDBrYFy12ygghlxCkYCF
# cwAso2lJl8CtBehInyLjUuTNR5cnEusL2xttI/mTdTcHjMdk1BA0efBJL575eQiF
# 7GLB3qXUbI+TSxIkreFF6/J9rW8Mv2riAGu3AeAzISXvkk1qZQ8aWwUEqBamT41Q
# MOA9RhG2EhOs9ou6r6STw6wEf4QhOngJTA33xz0r3jYB9KA=
# SIG # End signature block
