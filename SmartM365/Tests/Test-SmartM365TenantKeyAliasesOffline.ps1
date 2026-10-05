<#
.SYNOPSIS
Offline regression checks for profile selectors that differ from effective tenant keys.
.VERSION
1.0.0
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$sourceRoot = Split-Path $PSScriptRoot -Parent
$results = [System.Collections.Generic.List[object]]::new()

function Assert-TenantKeyCase {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}

function Test-TenantKeyCase {
    param([string]$Name, [scriptblock]$Body)
    try {
        & $Body
        $results.Add([pscustomobject]@{ Name = $Name; Passed = $true; Error = '' })
    }
    catch {
        $results.Add([pscustomobject]@{ Name = $Name; Passed = $false; Error = $_.Exception.Message })
    }
}

function Get-SourceFunction {
    param([string]$Path, [string]$Name)
    $tokens = $null
    $errors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors)
    if ($errors.Count -gt 0) { throw "Source parse failed: $Path" }
    $node = $ast.Find({
        param($candidate)
        $candidate -is [Management.Automation.Language.FunctionDefinitionAst] -and $candidate.Name -eq $Name
    }, $true)
    if (-not $node) { throw "Function '$Name' was not found in $Path" }
    return $node.Extent.Text
}

$winUpdatePath = Join-Path $sourceRoot 'SmartInventory/M365Inventory/IntuneInventory/WindowsUpdate/SmartM365-WinUpdate_Status_From_Intune.ps1'
$orchestratorPath = Join-Path $sourceRoot 'SmartInventory/Orchestrator/SmartM365-Inventory-Orchestrator.ps1'
$exchangePath = Join-Path $sourceRoot 'SmartInventory/ExchangeInventory/OnPremises/ServersAndStorage/SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory.ps1'
$winUpdateText = [IO.File]::ReadAllText($winUpdatePath)
$startupMatch = [regex]::Match(
    $winUpdateText,
    '(?m)^\$script:SmartM365GlobalConfig = Initialize-SmartM365TenantContext[^\r\n]*\r?\n\$CsvTenantKey = [^\r\n]*\r?\nif \([^\r\n]*\) \{ throw [^\r\n]* \}'
)
if (-not $startupMatch.Success) { throw 'WinUpdate effective TenantKey startup block was not found.' }
$startup = $startupMatch.Value
$savedTenantKey = Get-Variable -Name SmartM365TenantKey -Scope Global -ErrorAction SilentlyContinue

try {
    Test-TenantKeyCase 'WinUpdate -Tenant prod uses the effective example-prod key' {
        $stub = @'
$Tenant = 'prod'
function Initialize-SmartM365TenantContext {
    param([string]$Tenant, [string]$StartPath)
    if ($Tenant -ne 'prod') { throw 'Wrong profile selector.' }
    $global:SmartM365TenantKey = 'example-prod'
    return [pscustomobject]@{ ProfileKey = $Tenant; TenantKey = 'example-prod' }
}
'@
        $actual = & ([scriptblock]::Create($stub + [Environment]::NewLine + $startup + [Environment]::NewLine + '$CsvTenantKey'))
        Assert-TenantKeyCase ($actual -eq 'example-prod') 'WinUpdate used the profile selector instead of the effective tenant key.'
    }

    Test-TenantKeyCase 'WinUpdate rejects a missing effective tenant key' {
        $stub = @'
$Tenant = 'prod'
function Initialize-SmartM365TenantContext {
    param([string]$Tenant, [string]$StartPath)
    $global:SmartM365TenantKey = ''
    return [pscustomobject]@{ ProfileKey = $Tenant; TenantKey = '' }
}
'@
        $rejected = $false
        try { & ([scriptblock]::Create($stub + [Environment]::NewLine + $startup)) | Out-Null }
        catch { $rejected = $_.Exception.Message -match 'no TenantKey' }
        Assert-TenantKeyCase $rejected 'WinUpdate accepted a missing effective tenant key.'
    }

    Test-TenantKeyCase 'Orchestrator path and argument tokens separate profile and effective key' {
        $definitions = @(
            Get-SourceFunction -Path $orchestratorPath -Name 'Resolve-OrchestratorJobPath'
            Get-SourceFunction -Path $orchestratorPath -Name 'Resolve-OrchestratorJobArguments'
        ) -join [Environment]::NewLine
        $module = New-Module -ScriptBlock ([scriptblock]::Create($definitions))
        try {
            $root = Join-Path ([IO.Path]::GetTempPath()) 'SmartM365-TenantKey-Test'
            foreach ($case in @(
                @{ Profile = 'prod'; Effective = 'example-prod' },
                @{ Profile = 'test'; Effective = 'example-test' }
            )) {
                & $module { param($r, $p) $script:SmartInventoryRoot = $r; $script:Tenant = $p } $root $case.Profile
                $global:SmartM365TenantKey = $case.Effective
                $actual = & $module { Resolve-OrchestratorJobPath -Path 'Launchers\{{Tenant}}\{{TenantKey}}\run.cmd' }
                $expected = Join-Path $root ("Launchers\{0}\{1}\run.cmd" -f $case.Profile, $case.Effective)
                Assert-TenantKeyCase ($actual -eq $expected) "Orchestrator resolved the wrong path for $($case.Profile)."
                $arguments = & $module { Resolve-OrchestratorJobArguments -Arguments '-Profile {{Tenant}} -Key {{TenantKey}}' }
                Assert-TenantKeyCase ($arguments -eq ("-Profile {0} -Key {1}" -f $case.Profile, $case.Effective)) "Orchestrator resolved the wrong arguments for $($case.Profile)."
            }
            $global:SmartM365TenantKey = ''
            $rejected = $false
            try { & $module { Resolve-OrchestratorJobPath -Path 'Launchers\{{TenantKey}}\run.cmd' } | Out-Null }
            catch { $rejected = $_.Exception.Message -match 'effective tenant key' }
            Assert-TenantKeyCase $rejected 'Orchestrator accepted an unresolved effective tenant key.'
            $rejected = $false
            try { & $module { Resolve-OrchestratorJobArguments -Arguments '-Key {{TenantKey}}' } | Out-Null }
            catch { $rejected = $_.Exception.Message -match 'no configured value' }
            Assert-TenantKeyCase $rejected 'Orchestrator accepted an unresolved effective tenant key in arguments.'
            $profileOnly = & $module { Resolve-OrchestratorJobPath -Path 'Launchers\{{Tenant}}\run.cmd' }
            Assert-TenantKeyCase ($profileOnly -eq (Join-Path $root 'Launchers\test\run.cmd')) 'Orchestrator rejected a profile-only path.'
        }
        finally { Remove-Module $module -Force }
    }

    Test-TenantKeyCase 'Exchange completion log labels profile and effective key separately' {
        $definition = Get-SourceFunction -Path $exchangePath -Name 'Complete-ServersAndStorageRun'
        Assert-TenantKeyCase (
            $definition -match 'TenantProfile:\s*\{0\}"\s*-f\s*\$Tenant' -and
            $definition -match 'TenantKey:\s*\{0\}"\s*-f\s*\$global:SmartM365TenantKey'
        ) 'Exchange completion log still labels the profile selector as TenantKey.'
    }
}
finally {
    if ($savedTenantKey) { $global:SmartM365TenantKey = $savedTenantKey.Value }
    else { Remove-Variable -Name SmartM365TenantKey -Scope Global -ErrorAction SilentlyContinue }
}

$passed = @($results | Where-Object Passed).Count
$failed = $results.Count - $passed
[pscustomobject]@{ Passed = $passed; Failed = $failed; Total = $results.Count; Results = @($results) } | ConvertTo-Json -Depth 4
if ($failed -gt 0) { exit 1 }

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD1Epe3YUTntI/U
# 0zB/qKC12k57EEeQbfpORqPkQidieKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDetMxyD0VlqeTEdVzKr1LN
# sUnJBOeVNJLuCqMZ/cE/+zANBgkqhkiG9w0BAQEFAASCAYCjoHWyLz81VP33tS0L
# nlELxPAU+sQvV/fZCOFHiBndnXCivf/ghK4ua0IQZAZ4dXwsLr+NIChwLMiiVeRS
# ndhqa+tvNdHbxA/gr8bygVoKQpaXadPGpJhG/oOROvqrmAa1I8UC7GlqMASEr7Wj
# ecsblIEEqlNXgbZL8AsudnpXYIXxLa8l8avYRqW4KUiu0R/ibbu7ZoPOuxDU/4wg
# aq5ptHTkD02qJC7KKvtNqwIVBjRF4+RdXfI6WEsuqGioPZylx1RRPW6wPt9zTmOk
# hEQUsSwjXcWUy84T0XSHlIQtnqA+6nhpqn4n1OK2/ijWU8nSxe5/RO7uWl0GU6Lp
# 2gr+pGUNv+kR3vi/OO18nsCATqs2VHajkh1uy1ut5fb069GUNRTIiuGHXy6gLMMS
# 3pq98gcXr2NOz5nVUKm4txtiXSJQYr3jCIygI7pxp6V9SJSDjgDwydEzKLrTUnpB
# bhK4/A/aGlgGr6PNR8w7vo4y1FcPcBm0RS4j3fmggaSWnoE=
# SIG # End signature block
