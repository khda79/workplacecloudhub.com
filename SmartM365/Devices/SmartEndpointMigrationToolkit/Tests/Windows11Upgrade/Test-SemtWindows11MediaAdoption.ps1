<#
.SYNOPSIS
    Fault-injection tests of the reversible adoption of legacy Windows 11 setup media.
.DESCRIPTION
    Runs Invoke-SemtSetupMediaAdoption from the endpoint script on synthetic media in temporary
    folders owned by the test, with distinctive explicit permissions, real icacls save/reset/
    restore and real same-volume moves. Faults are injected before the move, during the move,
    during the permission reset, during the validation after the move and during the rollback.
    Every case checks paths, content hashes and permissions (DACL of every item) before/after,
    the reported outcome and the evidence file.
    Prints one PASS/FAIL line per assertion and "Summary: Passed=N; Failed=M". Exits 1 on failure.
.VERSION
    1.0.0
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

#requires -Version 5.1

[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$script:Passed = 0
$script:Failed = 0
$packageRoot = Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$endpointPath = Join-Path $packageRoot 'Campaigns\Windows11Upgrade\Scripts\SmartM365-Invoke-EndpointMigration-Windows11Upgrade.ps1'
$icacls = Join-Path $env:SystemRoot 'System32\icacls.exe'

function Write-TestResult {
    param([bool]$Condition, [string]$Message, [string]$Detail = '')
    if ($Condition) { $script:Passed++; Write-Host ("PASS: {0}" -f $Message) }
    else {
        $script:Failed++
        if ([string]::IsNullOrWhiteSpace($Detail)) { Write-Host ("FAIL: {0}" -f $Message) } else { Write-Host ("FAIL: {0} ({1})" -f $Message,$Detail) }
    }
}

function Test-Step {
    param([string]$Message, [scriptblock]$Check)
    try { Write-TestResult -Condition ([bool](& $Check)) -Message $Message }
    catch { Write-TestResult -Condition $false -Message $Message -Detail $_.Exception.Message }
}

function Get-RegionText {
    param([string]$Path, [string]$Region)
    $match = [regex]::Match([IO.File]::ReadAllText($Path), ('(?s)#region {0}.*?#endregion {0}' -f [regex]::Escape($Region)))
    if (-not $match.Success) { throw "Region '$Region' not found" }
    return $match.Value
}

foreach ($region in @('SEMT shared target guards v1','SEMT setup media resolution v1','SEMT legacy setup media reuse')) { . ([scriptblock]::Create((Get-RegionText -Path $endpointPath -Region $region))) }
$script:Log = New-Object System.Collections.Generic.List[string]
function Write-SmartLog { param([string]$Message, [string]$Level = 'INFO') $script:Log.Add(('{0} {1}' -f $Level,$Message)) }
$originals = @{}
foreach ($name in @('Save-SemtMediaAcl','Reset-SemtMediaAcl','Restore-SemtMediaAcl','Move-SemtMediaDirectory')) { $originals[$name] = (Get-Item -Path ('function:' + $name)).ScriptBlock }
function Reset-Overrides { foreach ($name in @($originals.Keys)) { Set-Item -Path ('function:script:' + $name) -Value $originals[$name] } }

$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('SEMT-AdoptTest-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
$noOwner = { param($Path) [pscustomobject]@{ State = 'Read'; Owners = @(); Detail = '' } }
$free = { param($Path) [pscustomobject]@{ State = 'Free'; Detail = '' } }
$script:ValidateFault = $null
$validator = {
    param($Path)
    if ($script:ValidateFault -and $Path -like '*\Data\*') { & $script:ValidateFault $Path }
    $ok = (Test-Path -LiteralPath (Join-Path $Path 'setup.exe')) -and (Test-Path -LiteralPath (Join-Path $Path 'sources\install.wim'))
    [pscustomobject]@{ Valid = $ok; Category = $(if ($ok) { 'Valid' } else { 'Corrupt' }); Detail = $(if ($ok) { '' } else { 'synthetic files missing' }) }
}

function New-AdoptionCase {
    param([string]$Name)
    $root = Join-Path $tempRoot $Name
    $legacy = Join-Path $root 'Legacy\SetupMedia\Win11-fr-FR'
    $cache = Join-Path $root 'Data\SetupMedia\Win11-fr-FR'
    New-Item -ItemType Directory -Path (Join-Path $legacy 'sources') -Force | Out-Null
    New-Item -ItemType Directory -Path (Split-Path $cache -Parent) -Force | Out-Null
    [IO.File]::WriteAllBytes((Join-Path $legacy 'setup.exe'), [Text.Encoding]::ASCII.GetBytes('SYNTHETIC-SETUP-' + ('x' * 2048)))
    [IO.File]::WriteAllBytes((Join-Path $legacy 'sources\install.wim'), [Text.Encoding]::ASCII.GetBytes('SYNTHETIC-WIM-' + ('y' * 4096)))
    # Distinctive permissions: explicit, protected DACL on the root plus an extra ACE on one file.
    $null = & $icacls $legacy '/inheritance:d' '/q' 2>&1
    $null = & $icacls $legacy '/grant' '*S-1-5-32-545:(OI)(CI)(RX)' '/q' 2>&1
    $null = & $icacls (Join-Path $legacy 'sources\install.wim') '/grant' '*S-1-1-0:(R)' '/q' 2>&1
    return [pscustomobject]@{ Root = $root; Legacy = $legacy; Cache = $cache; Evidence = (Join-Path $root 'Data\State\MediaAdoption') }
}

function Get-ContentState {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return '<absent>' }
    return (@(Get-ChildItem -LiteralPath $Path -Recurse -File | Sort-Object FullName | ForEach-Object { '{0}|{1}' -f $_.FullName.Substring($Path.Length),(Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash }) -join "`n")
}

function Get-AclState {
    param([string]$Path)
    return (@(Get-SemtMediaTreeSnapshot -Path $Path) | ForEach-Object { '{0}|{1}' -f $_.Path,(ConvertTo-SemtCanonicalDacl -Sddl $_.Sddl) }) -join "`n"
}

function Invoke-Adoption {
    param($Case, [scriptblock]$Owner = $noOwner, [scriptblock]$Usage = $free)
    return Invoke-SemtSetupMediaAdoption -LegacyPath $Case.Legacy -CachePath $Case.Cache -ValidateMedia $validator -EvidenceRoot $Case.Evidence -OwnerProvider $Owner -UsageProvider $Usage
}

function Get-Evidence {
    param($Result)
    return (Get-Content -LiteralPath (Join-Path $Result.EvidencePath 'adoption.json.txt') -Raw | ConvertFrom-Json)
}

try {
    #region Success
    $case = New-AdoptionCase 'success'
    $content = Get-ContentState $case.Legacy
    $result = Invoke-Adoption $case
    Test-Step 'success: outcome Moved, media at the new cache, legacy location empty' { $result.Outcome -eq 'Moved' -and $result.Path -eq $case.Cache -and -not (Test-Path -LiteralPath $case.Legacy) }
    Test-Step 'success: content identical after the move' { (Get-ContentState $case.Cache) -ceq $content }
    Test-Step 'success: permissions reset to inherited (no explicit or protected DACL left)' { @(@(Get-Item -LiteralPath $case.Cache) + @(Get-ChildItem -LiteralPath $case.Cache -Recurse) | Where-Object { $acl = Get-Acl -LiteralPath $_.FullName; $acl.AreAccessRulesProtected -or @($acl.GetAccessRules($true, $false, [System.Security.Principal.SecurityIdentifier])).Count -gt 0 }).Count -eq 0 }
    Test-Step 'success: evidence with saved permissions and every step' { $e = Get-Evidence $result; $e.Outcome -eq 'Moved' -and (Test-Path -LiteralPath (Join-Path $result.EvidencePath 'acl-before.icacls')) -and @($e.Steps | ForEach-Object { $_.Step }) -join ',' -eq 'PreChecks,SaveAcl,Move,ResetAcl,Validate' }
    #endregion

    #region Failures before the move
    foreach ($pre in @(
        @{ Name = 'target-exists'; Setup = { param($c) New-Item -ItemType Directory -Path $c.Cache -Force | Out-Null }; Owner = $noOwner; Usage = $free; Pattern = 'already exists' },
        @{ Name = 'owned'; Setup = { param($c) }; Owner = { param($Path) [pscustomobject]@{ State = 'Read'; Owners = @('Legacy-W11:Installed'); Detail = '' } }; Usage = $free; Pattern = 'owned by' },
        @{ Name = 'owner-unreadable'; Setup = { param($c) }; Owner = { param($Path) [pscustomobject]@{ State = 'Unreadable'; Owners = @(); Detail = 'x' } }; Usage = $free; Pattern = 'indeterminate' },
        @{ Name = 'in-use'; Setup = { param($c) }; Owner = $noOwner; Usage = { param($Path) [pscustomobject]@{ State = 'InUse'; Detail = 'Process setup.exe:1' } }; Pattern = 'not free' },
        @{ Name = 'invalid'; Setup = { param($c) Remove-Item -LiteralPath (Join-Path $c.Legacy 'sources\install.wim') -Force }; Owner = $noOwner; Usage = $free; Pattern = 'pre-move validation failed' },
        @{ Name = 'reparse'; Setup = { param($c) $target = Join-Path $c.Root 'junction-target'; New-Item -ItemType Directory -Path $target -Force | Out-Null; $null = & cmd.exe /d /c ('mklink /J "{0}" "{1}"' -f (Join-Path $c.Legacy 'sources\link'),$target) 2>&1 }; Owner = $noOwner; Usage = $free; Pattern = 'reparse point' }
    )) {
        $case = New-AdoptionCase ('pre-' + $pre.Name)
        & $pre.Setup $case
        $content = Get-ContentState $case.Legacy; $acl = Get-AclState $case.Legacy
        $result = Invoke-Adoption $case -Owner $pre.Owner -Usage $pre.Usage
        Write-TestResult -Condition ($result.Outcome -eq 'NotMoved' -and $result.Detail -match $pre.Pattern) -Message ("before move ({0}): NotMoved with reason" -f $pre.Name) -Detail $result.Detail
        Write-TestResult -Condition ((Get-ContentState $case.Legacy) -ceq $content -and (Get-AclState $case.Legacy) -ceq $acl) -Message ("before move ({0}): legacy content and permissions unchanged" -f $pre.Name)
        if ($pre.Name -eq 'reparse') { $null = & cmd.exe /d /c ('rmdir "{0}"' -f (Join-Path $case.Legacy 'sources\link')) 2>&1 }
    }
    $case = New-AdoptionCase 'pre-saveacl'
    Set-Item -Path 'function:script:Save-SemtMediaAcl' -Value { param($Path, $AclFile) [pscustomobject]@{ Success = $false; ExitCode = 5; Detail = 'injected' } }
    $content = Get-ContentState $case.Legacy; $acl = Get-AclState $case.Legacy
    $result = Invoke-Adoption $case
    Reset-Overrides
    Test-Step 'before move (permissions cannot be saved): NotMoved, nothing changed' { $result.Outcome -eq 'NotMoved' -and (Get-ContentState $case.Legacy) -ceq $content -and (Get-AclState $case.Legacy) -ceq $acl -and -not (Test-Path -LiteralPath $case.Cache) }
    #endregion

    #region Failures during the move
    $case = New-AdoptionCase 'move-error'
    Set-Item -Path 'function:script:Move-SemtMediaDirectory' -Value { param($Source, $Destination) throw 'injected: access denied before moving' }
    $content = Get-ContentState $case.Legacy; $acl = Get-AclState $case.Legacy
    $result = Invoke-Adoption $case
    Reset-Overrides
    Test-Step 'during move (error, nothing moved): NotMoved, content and permissions unchanged' { $result.Outcome -eq 'NotMoved' -and (Get-ContentState $case.Legacy) -ceq $content -and (Get-AclState $case.Legacy) -ceq $acl }

    $case = New-AdoptionCase 'move-then-error'
    Set-Item -Path 'function:script:Move-SemtMediaDirectory' -Value { param($Source, $Destination) [System.IO.Directory]::Move($Source, $Destination); if ($Destination -like '*\Data\*') { throw 'injected: error reported after the move' } }
    $content = Get-ContentState $case.Legacy; $acl = Get-AclState $case.Legacy
    $result = Invoke-Adoption $case
    Reset-Overrides
    Test-Step 'during move (moved, then error): RolledBack, content and permissions verified' { $result.Outcome -eq 'RolledBack' -and $result.AclRestored -and (Get-ContentState $case.Legacy) -ceq $content -and (Get-AclState $case.Legacy) -ceq $acl -and -not (Test-Path -LiteralPath $case.Cache) }
    #endregion

    #region Failures during the permission reset
    $case = New-AdoptionCase 'acl-partial'
    Set-Item -Path 'function:script:Reset-SemtMediaAcl' -Value { param($Path) $null = & (Join-Path $env:SystemRoot 'System32\icacls.exe') $Path '/reset' '/q' 2>&1; [pscustomobject]@{ Success = $false; ExitCode = 1332; Detail = 'injected: partial reset' } }
    $content = Get-ContentState $case.Legacy; $acl = Get-AclState $case.Legacy
    $result = Invoke-Adoption $case
    Reset-Overrides
    Test-Step 'during permission reset (partially applied, then error): RolledBack' { $result.Outcome -eq 'RolledBack' -and $result.AclRestored }
    Test-Step 'during permission reset: original permissions restored exactly, content intact' { (Get-AclState $case.Legacy) -ceq $acl -and (Get-ContentState $case.Legacy) -ceq $content -and -not (Test-Path -LiteralPath $case.Cache) }
    #endregion

    #region Failure during the validation after the move (after real permission changes)
    $case = New-AdoptionCase 'validate'
    $script:ValidateFault = { param($Path) throw 'injected: media invalid after move' }
    $content = Get-ContentState $case.Legacy; $acl = Get-AclState $case.Legacy
    $result = Invoke-Adoption $case
    $script:ValidateFault = $null
    Test-Step 'after-move validation failure: RolledBack after the real permission reset' { $result.Outcome -eq 'RolledBack' -and @((Get-Evidence $result).Steps | Where-Object { $_.Step -eq 'ResetAcl' -and $_.Status -eq 'OK' }).Count -eq 1 }
    Test-Step 'after-move validation failure: location, content and permissions identical to the initial state' { (Get-AclState $case.Legacy) -ceq $acl -and (Get-ContentState $case.Legacy) -ceq $content -and -not (Test-Path -LiteralPath $case.Cache) }
    Test-Step 'after-move validation failure: evidence keeps both snapshots and the failure' { (Test-Path -LiteralPath (Join-Path $result.EvidencePath 'snapshot-before.json.txt')) -and (Test-Path -LiteralPath (Join-Path $result.EvidencePath 'snapshot-after-rollback.json.txt')) -and ((Get-Evidence $result).Detail -match 'validation after move failed') }
    #endregion

    #region Failures during the rollback
    $case = New-AdoptionCase 'rollback-move'
    $script:ValidateFault = { param($Path) throw 'injected: media invalid after move' }
    Set-Item -Path 'function:script:Move-SemtMediaDirectory' -Value { param($Source, $Destination) if ($Destination -like '*\Legacy\*') { throw 'injected: move back denied' }; [System.IO.Directory]::Move($Source, $Destination) }
    $content = Get-ContentState $case.Legacy
    $result = Invoke-Adoption $case
    Reset-Overrides; $script:ValidateFault = $null
    Test-Step 'rollback move failure: RollbackFailed, reported location is where the media really is' { $result.Outcome -eq 'RollbackFailed' -and $result.Path -eq $case.Cache -and (Test-Path -LiteralPath (Join-Path $case.Cache 'setup.exe')) }
    Test-Step 'rollback move failure: no data lost (content intact at the reported location)' { (Get-ContentState $case.Cache) -ceq $content }

    $case = New-AdoptionCase 'rollback-acl'
    $script:ValidateFault = { param($Path) throw 'injected: media invalid after move' }
    Set-Item -Path 'function:script:Restore-SemtMediaAcl' -Value { param($Path, $Snapshot) [pscustomobject]@{ Success = $false; ExitCode = 5; Detail = 'injected: restore denied' } }
    $content = Get-ContentState $case.Legacy; $acl = Get-AclState $case.Legacy
    $result = Invoke-Adoption $case
    Reset-Overrides; $script:ValidateFault = $null
    Test-Step 'rollback permission restore failure: RollbackPartial (not announced as complete)' { $result.Outcome -eq 'RollbackPartial' -and -not $result.AclRestored -and $result.Detail -match 'not complete' }
    Test-Step 'rollback permission restore failure: media back at the legacy location, content intact, permissions really differ' { $result.Path -eq $case.Legacy -and (Get-ContentState $case.Legacy) -ceq $content -and (Get-AclState $case.Legacy) -cne $acl }

    $case = New-AdoptionCase 'rollback-recreated'
    $script:ValidateFault = { param($Path) $legacyPath = $Path.Replace('\Data\', '\Legacy\'); New-Item -ItemType Directory -Path $legacyPath -Force | Out-Null; Set-Content -LiteralPath (Join-Path $legacyPath 'concurrent.txt') -Value 'recreated' -Encoding ASCII; throw 'injected: media invalid after move' }
    $content = Get-ContentState $case.Legacy
    $result = Invoke-Adoption $case
    $script:ValidateFault = $null
    Test-Step 'legacy folder recreated concurrently: RollbackFailed, the recreated folder is not overwritten' { $result.Outcome -eq 'RollbackFailed' -and (Get-Content -LiteralPath (Join-Path $case.Legacy 'concurrent.txt') -Raw).Trim() -eq 'recreated' -and @(Get-ChildItem -LiteralPath $case.Legacy -Recurse -File).Count -eq 1 }
    Test-Step 'legacy folder recreated concurrently: media kept intact at the new cache' { $result.Path -eq $case.Cache -and (Get-ContentState $case.Cache) -ceq $content }
    #endregion
}
catch {
    Write-TestResult -Condition $false -Message 'test harness' -Detail ($_.Exception.Message + ' @ ' + $_.InvocationInfo.PositionMessage)
}
finally {
    Reset-Overrides
    foreach ($dir in @(Get-ChildItem -LiteralPath $tempRoot -Recurse -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Attributes -band [IO.FileAttributes]::ReparsePoint })) { $null = & cmd.exe /d /c ('rmdir "{0}"' -f $dir.FullName) 2>&1 }
    $null = & $icacls $tempRoot '/reset' '/t' '/c' '/q' 2>&1
    try { [IO.Directory]::Delete('\\?\' + $tempRoot, $true) } catch { Write-Host ("WARN: temp folder not fully removed: {0}" -f $tempRoot) }
}

Write-Host ("Summary: Passed={0}; Failed={1}" -f $script:Passed,$script:Failed)
if ($script:Failed -gt 0) { exit 1 }
exit 0

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCbxzPIjdcYGHNH
# zcEGtT6h+PZuNELUZO9d5tMBcTaiNKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCGDAio/G5RvfvYhs9cKcMV
# FGyYDsHFp2JI75GGmmksHTANBgkqhkiG9w0BAQEFAASCAYBm+Zh1L9Y0PnMh1sqP
# m8V+K7Y0uOOxUuWOgLJxdu2YGWmSngWy9pGRehkIGc6XBbfEDOCG1ZtRFlRInte4
# Rd1NWaV2VocLu14OOy2tmqeCip2Q5p5KiMQHNiLQdEHIH/5VjvRCHKSWE2fo/uzB
# uLTnw4YGjJSa0akc2UOyRPaVyqSO5RVINbZl7vmEhaMXJYSwWD0dFWldwzgY3TK8
# jVbD+WBr54TPHpa3p9yzS50HqWkv9eLLCoXZPzoKA5dzfqja/xKaA1FdzMWlGcFv
# QATDKEjoXGa0VB00XfMGRdiRi68J4fsWXw0MQZ4+N8MWs3vlCzgbJ4aVpPWL5kze
# 6nR7qBGP0kQtOKe9ix+GLWs8wLsPSKDksLaWTRpQDTE76WLzkVoK7nV/41c/fYf7
# gveMDwodFmuieUyWID76eeRIaXAVovpIp6vJqsj3+/SVLYk65Qx9R5BAIKwqaYCK
# 5HaiBef14I9BJX4SWyDdt/kAT0kfZWf+ycUbjXl/oMM5Y0g=
# SIG # End signature block
