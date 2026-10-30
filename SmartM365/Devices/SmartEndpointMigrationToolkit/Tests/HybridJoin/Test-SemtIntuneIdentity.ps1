<#
.VERSION
1.0.0
#>
#requires -Version 5.1
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
Import-Module (Join-Path $root 'Modules\SmartEndpointMigration.Core\SmartEndpointMigration.Core.psd1') -MinimumVersion 1.0.0 -Force
Import-Module (Join-Path $root 'Campaigns\HybridJoin\SmartEndpointMigration.Campaign.HybridJoin.psd1') -MinimumVersion 1.0.0 -Force
$module = Get-Module SmartEndpointMigration.Campaign.HybridJoin
& $module {
    $id = '11111111-1111-1111-1111-111111111111'
    $tenant = '22222222-2222-2222-2222-222222222222'
    $row = [pscustomobject]@{ Dsreg_DeviceId=$id; Dsreg_TenantId=$tenant; IntuneEnrolled=$true }
    $cloud = [pscustomobject]@{ AzureADDeviceId=$id; InventoryTenantId=$tenant; DeviceName='OTHER-NAME' }
    function Assert-State($actual,$expected) { if ($actual -ne $expected) { throw "Expected $expected, got $actual" }; Write-Host "PASS: $expected" }
    Assert-State (Get-HybridJoinIntuneConfirmation -Row $row -CloudRows @($cloud)).State 'Confirmed'
    $wrong = [pscustomobject]@{ AzureADDeviceId='33333333-3333-3333-3333-333333333333'; InventoryTenantId=$tenant; DeviceName='SAME-NAME' }
    Assert-State (Get-HybridJoinIntuneConfirmation -Row $row -CloudRows @($wrong)).State 'NotFound'
    Assert-State (Get-HybridJoinIntuneConfirmation -Row $row -CloudRows @($cloud,$cloud)).State 'Ambiguous'
    $row.IntuneEnrolled = $false
    Assert-State (Get-HybridJoinIntuneConfirmation -Row $row -CloudRows @($cloud)).State 'CloudOnly'
    $row.IntuneEnrolled = $true
    $cloud.InventoryTenantId = '44444444-4444-4444-4444-444444444444'
    Assert-State (Get-HybridJoinIntuneConfirmation -Row $row -CloudRows @($cloud)).State 'Unknown'
    $row.Dsreg_DeviceId = ''
    Assert-State (Get-HybridJoinIntuneConfirmation -Row $row -CloudRows @($cloud)).State 'Unknown'
    if (Test-HybridJoinAlreadyEnrolled ([pscustomobject]@{ IntuneInventoryPresent=$true; Status='SUCCESS' })) { throw 'Name-only result was archived.' }
    if (Test-HybridJoinAlreadyEnrolled ([pscustomobject]@{ IntuneConfirmationState='Confirmed'; Status='GPUPDATE_TIMEOUT' })) { throw 'A failed run was archived.' }
    Write-Host 'PASS: name-only result is not archived'
}

# Preflight never treats missing or ambiguous evidence as a license absence.
$probe=[pscustomobject]@{ Sid='S-1-5-21-42'; Tenants=@('22222222-2222-2222-2222-222222222222'); AutoEnroll=1; CredentialType=1 }
$user=[pscustomobject]@{ InventoryTenantId=$probe.Tenants[0]; InventoryScope='AllUsers'; InventoryComplete='True'; OnPremisesSecurityIdentifier=$probe.Sid; AssignedSkuIds='' }
if (-not (Get-SemtUserLicensePreflightDecision -Evidence $probe -Users @($user)).Block) { throw 'Confirmed unlicensed user was not guarded.' }
$user.AssignedSkuIds='sku'
if ((Get-SemtUserLicensePreflightDecision -Evidence $probe -Users @($user)).Block) { throw 'User with an assigned SKU was incorrectly blocked.' }
$user.AssignedSkuIds=''
if ((Get-SemtUserLicensePreflightDecision -Evidence $probe -Users @($user,$user)).Block) { throw 'Ambiguous SID was incorrectly blocked.' }
$probe.CredentialType=2
if ((Get-SemtUserLicensePreflightDecision -Evidence $probe -Users @($user)).Block) { throw 'Device credential enrollment was incorrectly blocked.' }
Write-Host 'PASS: preflight guards no-SKU evidence and leaves ambiguous cases unknown'

# Exercise the target attempt contract, not just the technician-side classifier.
$endpoint = Join-Path $root 'Campaigns\HybridJoin\Scripts\SmartM365-Invoke-EndpointMigration-HybridJoin.ps1'
$tokens=$null; $errors=$null
$ast=[System.Management.Automation.Language.Parser]::ParseFile($endpoint,[ref]$tokens,[ref]$errors)
if ($errors.Count) { throw 'Endpoint parse failed.' }
$definition=$ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Write-SemtHybridJoinAttemptResult' },$true)
. ([scriptblock]::Create($definition.Extent.Text))
function Write-SemtAttemptResult { param($CampaignDataRoot,$Data) return $Data }
$script:SemtCampaign='HybridJoin'; $script:SemtAttemptStartUtc=[datetime]::UtcNow
$script:SemtCsvPath=''; $script:SemtTranscriptPath=''; $script:SemtTargetGuardStatus=''; $script:SemtLegacyActivityStatus=''; $script:SemtLegacyActivityDetail=''
$script:RetryAfterRebootAction=''; $script:RetryAfterRebootDetail=''; $script:RetryAfterRebootAttempt=''; $script:RetryAfterRebootMaxAttemptsResult=''; $script:RetryAfterRebootTaskNameResult=''
$ScriptVersion='1.0.0'; $RunId='test'; $CorrelationId='test'; $ComputerName='PC'; $RunLogPath=''; $DataRoot=''; $RetryAfterRebootTaskRun=$false
$script:dsregDeviceId='11111111-1111-1111-1111-111111111111'; $script:dsregTenantId='22222222-2222-2222-2222-222222222222'; $script:interactiveUserSid='S-1-5-21-1'; $script:mdmEnrollmentRecentEvidence='event'
$attempt=Write-SemtHybridJoinAttemptResult -State Completed -Status SUCCESS
if ($attempt.Dsreg_DeviceId -ne $script:dsregDeviceId -or $attempt.Dsreg_TenantId -ne $script:dsregTenantId -or $attempt.InteractiveUserSid -ne $script:interactiveUserSid -or $attempt.MdmEnrollmentRecentEvidence -ne 'event') { throw 'Attempt identity evidence was not projected.' }
Write-Host 'PASS: endpoint attempt carries identity and MDM evidence'

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCsPocvZVVILU8u
# Iiy0RHfZWtx0W0YIlzJTIVWTyGRqFqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDx/dCTdsdKapFadF32OFjk
# YGL4/veXymhb0JwMvL6sMzANBgkqhkiG9w0BAQEFAASCAYBChiGa7fgUlFXqmg3q
# bJzz5R/WdGoXyhykJLtyC1ZSykXEwhpuo/RvCTfxJjD6KkbgmWeLl2jLaO87V2F6
# LgbPCOZMeyDJwIr+r/mlfmN7jlOlZkm/DBAKAYUpn1Rj7cMli5R/pEZcBp0iKTpV
# erbBBlqxGXioEugZAGW73NE5LK14KSnSj6NxMxxPqGyKoYKscG0jM4Mx7oHe4RF5
# L1pUe/Br0RON4/0Md6Wj7HdoM9fi608gWmgo9KJRs0AILafFKEcR7ftjH6q4zU3Q
# aGuZ68dPfnOxAFSbWhZYuKRrAUjEvr+oFNx5Lf1Wdol/n+whawHKbdECfnqoKnVi
# T725MFKsNSPvhCQ04oSE3ktPwDv+KRv1NnFj4HdxnEDQoZUHDqPC7UhOTvQISNDm
# ZWDQP8ij/qB1NV3DPIUH63ALVRew11n9bfon/TdQHvOg+jej8VaudYEYq46xavq+
# MOSAvVHPHKilPa8dNClZfEkfx8Mn0UkQMnuFYGvm4BcHP3k=
# SIG # End signature block
