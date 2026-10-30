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
$user=[pscustomobject]@{ InventoryTenantId=$probe.Tenants[0]; InventoryScope='AllUsers'; InventoryComplete='True'; OnPremisesSecurityIdentifier=$probe.Sid; UserPrincipalName='user@example.test'; AssignedSkuIds=''; CollectedUtc=[datetime]::UtcNow.ToString('o') }
$decision=Get-SemtUserLicensePreflightDecision -Evidence $probe -Users @($user)
if (-not $decision.Block -or $decision.UserPrincipalName -ne 'user@example.test' -or $decision.CollectedUtc -ne $user.CollectedUtc) { throw 'Confirmed unlicensed user was not guarded or identified.' }
$user.CollectedUtc=[datetime]::UtcNow.AddHours(-25).ToString('o')
if ((Get-SemtUserLicensePreflightDecision -Evidence $probe -Users @($user)).Block) { throw 'Stale user evidence was treated as an absent license.' }
$user.CollectedUtc='2026-10-30T21:00:00'
if ((Get-SemtUserLicensePreflightDecision -Evidence $probe -Users @($user)).Block) { throw 'Timezone-unspecified evidence was treated as an absent license.' }
$user.CollectedUtc=[datetime]::UtcNow.AddHours(1).ToString('o')
if ((Get-SemtUserLicensePreflightDecision -Evidence $probe -Users @($user)).Block) { throw 'Future user evidence was treated as an absent license.' }
$user.CollectedUtc=[datetime]::UtcNow.ToString('o')
$missingColumn=[pscustomobject]@{ InventoryTenantId=$probe.Tenants[0]; InventoryScope='AllUsers'; InventoryComplete='True'; OnPremisesSecurityIdentifier=$probe.Sid; CollectedUtc=$user.CollectedUtc }
if ((Get-SemtUserLicensePreflightDecision -Evidence $probe -Users @($missingColumn)).Block) { throw 'Missing license field was treated as an absent license.' }
$user.AssignedSkuIds=$null
if ((Get-SemtUserLicensePreflightDecision -Evidence $probe -Users @($user)).Block) { throw 'Null license field was treated as an absent license.' }
$user.AssignedSkuIds='sku'
if ((Get-SemtUserLicensePreflightDecision -Evidence $probe -Users @($user)).Block) { throw 'User with an assigned SKU was incorrectly blocked.' }
$user.AssignedSkuIds=''
if ((Get-SemtUserLicensePreflightDecision -Evidence $probe -Users @($user,$user)).Block) { throw 'Ambiguous SID was incorrectly blocked.' }
$probe.CredentialType=2
if (-not (Get-SemtUserLicensePreflightDecision -Evidence $probe -Users @($user)).Block) { throw 'An unlicensed user was not guarded for device credential enrollment.' }
$probeWithoutPolicy=[pscustomobject]@{ Sid=$probe.Sid; Tenants=$probe.Tenants }
if (-not (Get-SemtUserLicensePreflightDecision -Evidence $probeWithoutPolicy -Users @($user)).Block) { throw 'An unlicensed user was not guarded without MDM policy data.' }
Write-Host 'PASS: preflight guards no-SKU evidence and leaves ambiguous cases unknown'

# Both audit and repair check licenses before copying or starting the endpoint.
$worker = Join-Path $root 'Scripts\SmartM365-EndpointMigration-PsExecWorker.ps1'
$tokens=$null; $errors=$null
$workerAst=[System.Management.Automation.Language.Parser]::ParseFile($worker,[ref]$tokens,[ref]$errors)
if ($errors.Count) { throw 'Worker parse failed.' }
$preflightIf=$workerAst.Find({param($n) $n -is [System.Management.Automation.Language.IfStatementAst] -and $n.Extent.Text.StartsWith("if (`$campaign -eq 'HybridJoin'") -and $n.Extent.Text.Contains('Invoke-SemtUserLicensePreflight')},$true)
if (-not $preflightIf) { throw 'Worker license preflight branch missing.' }
$runPreflight=[scriptblock]::Create($preflightIf.Extent.Text)
function Invoke-SemtUserLicensePreflight { param($ComputerName,$InventoryPath); return [pscustomobject]@{State='NoAssignedLicense';Detail='No user SKU';Block=$true} }
function Write-SemtLogLine { param($Path,$Message) }
$campaign='HybridJoin'; $target='fixture.invalid'; $packageRoot=$root; $logPath='unused'; $access=[pscustomobject]@{AdminShareReachable=$true}
foreach ($arguments in @(@('-AuditOnly'),@('-RepairFixture'))) {
    $contract=[pscustomobject]@{Flags=[pscustomobject]@{DryRun=$false}; Endpoint=[pscustomobject]@{Arguments=@($arguments)}}
    $result=@{}
    & $runPreflight
    if ($result.UserLicensePreflightState -ne 'NoAssignedLicense') { throw "Preflight omitted for $arguments" }
}
Write-Host 'PASS: audit and repair both run the license preflight'
$workerDecision=$workerAst.Find({param($n) $n -is [System.Management.Automation.Language.IfStatementAst] -and $n.Extent.Text.StartsWith('if ([bool]$contract.Flags.DryRun)') -and $n.Extent.Text.Contains('USER_LICENSE_REVIEW_REQUIRED')},$true)
if (-not $workerDecision -or $workerDecision.Clauses.Count -lt 2) { throw 'Worker license-stop branch missing.' }
$clause=$workerDecision.Clauses[1]
$runStop=[scriptblock]::Create('if (' + $clause.Item1.Extent.Text + ') ' + $clause.Item2.Extent.Text)
function Set-WorkerStatus { param($Status,$NextAction,$Detail); $script:observedLicenseStatus=$Status; $script:observedLicenseAction=$NextAction }
$licensePreflight=[pscustomobject]@{Block=$true;Detail='No user license'}
& $runStop
if ($script:observedLicenseStatus -ne 'USER_LICENSE_REVIEW_REQUIRED' -or $script:observedLicenseAction -ne 'ASSIGN_INTUNE_LICENSE_THEN_RETRY') { throw 'Unlicensed user did not receive the stop status and corrective action.' }
Write-Host 'PASS: a confirmed unlicensed user receives the stop status and license action'

# A blocked computer is not retried during a Loop Run or removed from its Wave.
$enginePath=Join-Path $root 'Scripts\SmartM365-Invoke-EndpointMigrationWithPsExec.ps1'
$tokens=$null; $errors=$null
$engineAst=[System.Management.Automation.Language.Parser]::ParseFile($enginePath,[ref]$tokens,[ref]$errors)
if ($errors.Count) { throw 'Engine parse failed.' }
$filterIf=$engineAst.Find({param($n) $n -is [System.Management.Automation.Language.IfStatementAst] -and $n.Extent.Text.StartsWith("if (`$Campaign -eq 'HybridJoin' -and `$licenseBlockedThisRun.Count")},$true)
if (-not $filterIf) { throw 'License-blocked Loop Run filter missing.' }
$runFilter=[scriptblock]::Create('while ($true) { ' + $filterIf.Extent.Text + '; break }; [pscustomobject]@{Computers=@($computers); Reason=$stopReason}')
$Campaign='HybridJoin'; $licenseBlockedThisRun=@{ 'PC-UNLICENSED' = $true }
$computers=@('PC-UNLICENSED','PC-LICENSED'); $stopReason='Finished'; $filtered=& $runFilter
if (@($filtered.Computers).Count -ne 1 -or $filtered.Computers[0] -ne 'PC-LICENSED') { throw 'Loop Run retried a license-blocked computer.' }
$computers=@('PC-UNLICENSED'); $stopReason='Finished'; $filtered=& $runFilter
if ($filtered.Reason -ne 'BlockedUserLicense' -or @($filtered.Computers).Count -ne 0) { throw 'Loop Run did not stop when all remaining computers are license-blocked.' }
$licenseBlockedThisRun=@{ 'PC1.FIRST.EXAMPLE' = $true }
$computers=@('pc1.first.example','pc1.second.example'); $stopReason='Finished'; $filtered=& $runFilter
if (@($filtered.Computers).Count -ne 1 -or $filtered.Computers[0] -ne 'pc1.second.example') { throw 'Distinct domain names were incorrectly treated as the same blocked computer.' }
Write-Host 'PASS: Loop Run leaves unlicensed computers pending without retrying them'

# A stopped unlicensed device is reported, and does not trigger post-cycle Graph calls.
& $module {
    $row=[pscustomobject]@{Status='USER_LICENSE_REVIEW_REQUIRED'; UserLicensePreflightState='NoAssignedLicense'; UserLicensePreflightDetail='No user SKU'; InteractiveUserSid=''; Dsreg_TenantId=''}
    $context=[pscustomobject]@{PackageRoot=(Join-Path ([IO.Path]::GetTempPath()) ('SEMT-absent-'+[guid]::NewGuid().ToString('N'))); DryRun=$false; Log={param($text,$color)}; CycleNumber=1}
    Invoke-HybridJoinPostCycle -Context $context -Rows @($row)
    if ($row.Status -ne 'USER_LICENSE_REVIEW_REQUIRED' -or $row.UserLicenseState -ne 'NoAssignedLicense' -or $row.UserLicenseDetail -ne 'No user SKU') { throw 'License-blocked result was not preserved.' }
    if ($row.PSObject.Properties['PostCycleIntuneInventoryChecked']) { throw 'License-blocked device triggered cloud confirmation.' }
}
Write-Host 'PASS: license-blocked device remains pending and skips post-cycle Graph confirmation'

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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDJ4ZT2BI4WqP5g
# wmzHUaqP4IQ7Dwl6bdZrHGbMx4i09KCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAHqimwaDS3lh5wCjtXnYma
# buu58VpT9wdKSmxi3wzmsjANBgkqhkiG9w0BAQEFAASCAYBw1rTOm75cAbS25Nd9
# iCLdWJlsOxevskd7jktGzeQIQtieLLvfRxyZ8u7Vedfx140YwxccImYkRbUvNhPj
# L5RRMTjgyY5Fi2cyJ2sSGeSxukIRPDNexKwn3drOfNh1AouX0eL9vFUe2h0twFKm
# GHfp7yU8TVaHnJz172z2YBXLZzjpjYbITiHAIzG7GkV4GseHrKV5wvMGwqzeCxgf
# ddM+11w+yTuj2VNjWb0EAxW0suEW5i7k6anVIAxkB5UGpfmbFLijyI8mO/6ylB13
# KoYgrsb2Rxypl7oFGzpla4RIP1rNvuivCcM4I2VLodQ1KPIzXHlcvO3gze70Raj7
# FHrf5pNEOr29TJtS/3ms49LJ1CevKn1q+LQpiuA7bl3DV7FyBHPukCRfcGcKZD2w
# ERwhs/14Jx79mLTpYJcxnnJP5ZiyLzaqji1h6BHfTnU6IZmW3WwnibdMn5604x9I
# qzW1y6UIuUgImb0LCFzqASGppsSRnfJWGigDc/i0332jXn4=
# SIG # End signature block
