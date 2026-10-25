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
    return (Get-Content -LiteralPath (Join-Path $Result.EvidencePath 'adoption.json') -Raw | ConvertFrom-Json)
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
    Test-Step 'after-move validation failure: evidence keeps both snapshots and the failure' { (Test-Path -LiteralPath (Join-Path $result.EvidencePath 'snapshot-before.json')) -and (Test-Path -LiteralPath (Join-Path $result.EvidencePath 'snapshot-after-rollback.json')) -and ((Get-Evidence $result).Detail -match 'validation after move failed') }
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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBQKNVjCw6t/SSb
# qdeVvAcYndIjZX5QmOj4EahV9cZl96CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIHWPdyJWVQP4UJFaY1slwQIS8snlzFBEQlubrDytGD3HMA0GCSqG
# SIb3DQEBAQUABIIBgHmleDX4tdjGcGPOoJlPrizf6zqNuMwDa/wIoznRfu2cZnXl
# tkAMDOkdQBl6vfWWHRubZSpoD0XQHuDdPZGFjph7Sik6dKqb5foy+L5gsaT2PdLD
# lW8ZjeUXa5JFsy+k68Cf1IaiTS1+70B5c6/+5/2KuWf501dDPUjgxoh8mBo8uQgt
# m+b51FnSYJo+CXnb00DWGJ6mx2587hS+anLAmArORRe/GASYFmQKqcGhSbSWpLlm
# sk7FvVy7Bs10cSx6QXkIraAPZbR0u9mjNHuz5BOV8G8OFrwyXCwTa/3KCYP5H7vs
# 548wPOd6ZygqjTJz3cBCWZksRxTHyQVODFt6YtAYnePDb6+qTxwkFedv93IJ28sy
# aHudGXtu2mA1ql8VlGgYAI7VzKJnHNaIUJTxpwoOFGNpCD4cVlhbJirIPp+lIU8d
# YOZIwa4ELFGKyyFaSffOcQ5y0LM4o69R6OaIwYsEs3wZhxMqoVDS/P3/9f0YLqKp
# uD+gX0Ch6Cghf0U3h6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjUxMzE2
# NTlaMC8GCSqGSIb3DQEJBDEiBCCQEn8IGnLsJPc8k1tKldUytuyXJ6xpMPJ64dSE
# iZI3vTANBgkqhkiG9w0BAQEFAASCAgBGb+Sdg2VzIDKzNW0UwIl8YGgVLlOup4rY
# KjQuLqRLRp/jkfbOfv35vDI2lhfUVhvycY6afpPTrcAwVk77rpsDKKZGQFVbTfEY
# 9f29garp7T5EsneL8HcFFoeFNPiv7jNQcuckS+foBG85630sEDhfFsZBQa3DlnbA
# z425JtJ2VFGiQ3DI0hGRROFjJ51j9QVedFtzefx79kcMEv+avxUQLhLXIgYyp4bn
# NtJ+eY6OA6KtCKAlNzn9+9Z1O7xASd8x7lY8AXfsaWZdASNon9f4McdRLg7YNCSH
# o4VFlgYrp7F6VcAnjBK4JKas/Q+GARqlVJtATkZ1nB16Rpl8VBnVc07woTcApNm9
# WkgTwyc5dcQfEKRc+4QWG7zTpiLr3jXLUwWZRAkibpPHVIX48xZm6B6sGIjLCPLw
# I5PeW1gThEMPWv5/d8KPccAeBdfvjQ1I/juMNXH634+TQFmwGn6i1iIDNCgHi524
# N68k5L/GVahxSGWWVLBs97ER2wjFn5OdWR5BFkGT/of8Fubpsi+LqqxQtQKkZlys
# j9F+RSootYAQ54vBbdpOE7tVg7uIFJ5p4BEAja0W9J0co1G+xxOxJImkmOYy6ACa
# 1eDmEWB/c+NsA0iAocZeYxN16H31A1zUSmrm5A74JemsMGjSbK2k4Hf/bEijTYra
# TnYQEHuVSg==
# SIG # End signature block
