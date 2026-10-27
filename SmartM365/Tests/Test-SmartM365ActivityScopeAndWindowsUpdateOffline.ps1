# Version 1.1.0
<# Synthetic scope and temporal-selection consumer regressions. No operational entry point. #>
[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot '../SmartInventory/Common/SmartM365.IdentityDateEvidence.ps1')
$checks=0
function Check([bool]$Ok,[string]$Message){if(-not $Ok){throw $Message};$script:checks++}
function LoadAst([string]$Relative){$t=$null;$e=$null;$a=[Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot "../$Relative"),[ref]$t,[ref]$e);Check ($e.Count -eq 0) "Parser: $Relative";return $a}
$users=LoadAst 'SmartInventory/ActiveDirectoryInventory/SmartM365-ActiveDirectory-UsersEnrichment.ps1'
$pc=LoadAst 'SmartInventory/ActiveDirectoryInventory/SmartM365-ActiveDirectory-Enrichment.ps1'
$now=[datetimeoffset]'2026-10-22T12:00:00Z'
$old=[pscustomobject]@{Status='Qualified';Utc=[datetimeoffset]'2025-01-01T00:00:00Z';Proof='Synthetic'}
$missing=Read-IdentityDate $null LastLogonTime
$na=[pscustomobject]@{Status='Not applicable';Utc=$null;Proof='Positive synthetic scope proof'}
Check ((Merge-IdentityActivity @($old,$na) 180 -Now $now).Old -eq $true) 'Excluded source dilutes applicable scope'
Check ($null -eq (Merge-IdentityActivity @($old,$missing) 180 -Now $now).Old) 'Missing expected source excluded'
Check ($null -eq (Merge-IdentityActivity @($na) 180 -Now $now).Old) 'No applicable source implies inactivity'
$guid='11111111-1111-1111-1111-111111111111'
$ad=[pscustomobject]@{ObjectGUID=$guid;LastLogonDateIso8601='2025-01-01T00:00:00Z'}
$remote=[pscustomobject]@{ObjectGUID=$guid;RecipientTypeDetails='RemoteUserMailbox'}
$cloud=[pscustomobject]@{LastLogonTimeIso8601='2025-01-01T00:00:00Z'}
$recent=[pscustomobject]@{LastLogonTimeIso8601='2026-10-21T00:00:00Z'}
foreach($consumer in @($users)) {
    $variable=if($consumer -eq $users){'$mailboxActivitySet'}else{'$activityEvidence'}
    $assignment=$consumer.Find({param($n)$n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq $variable},$true)
    Check ($null -ne $assignment) 'Actual consumer list not found'
    foreach($scenario in @('remote-old','expected-missing','collection-error','conflict','duplicate','recent-partial')) {
        $user=$ad;$u=$ad;$localMailbox=$null;$lm=$null;$exoMailbox=$cloud;$em=$cloud;$exoStat=$null;$st=$null
        $remoteRows=@($remote)
        switch($scenario){
            'expected-missing' {$remoteRows=@()}
            'collection-error' {$remoteRows=@();$localMailbox=[pscustomobject]@{LastLogonTime='Error'};$lm=$localMailbox}
            'conflict' {$localMailbox=$cloud;$lm=$cloud}
            'duplicate' {$remoteRows=@($remote,$remote)}
            'recent-partial' {$remoteRows=@();$exoMailbox=$recent;$em=$recent}
        }
        $remoteScopeByGuid=New-IdentityRemoteScopeIndex $remoteRows
        . ([scriptblock]::Create($assignment.Extent.Text))
        $entries=if($consumer -eq $users){$mailboxActivitySet}else{$activityEvidence}
        $merged=Merge-IdentityActivity $entries 180 -Now $now
        if($scenario -eq 'remote-old') {Check ($merged.Old -eq $true -and $merged.ExcludedCount -eq 1 -and $merged.Exclusions) 'Native remote scope did not remove only local logon'}
        elseif($scenario -eq 'recent-partial'){Check ($merged.Recent -eq $true -and $merged.Coverage -eq 'Partial') 'Positive recent cloud activity lost'}
        else {Check ($null -eq $merged.Old -and $merged.ExcludedCount -eq 0) "Unknown/ambiguous scope excluded: $scenario"}
    }
}
# Execute the actual device-list assignment, with missing native maps preserved as unknown.
$deviceAssignment=$pc.Find({param($n)$n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$deviceActivity' -and $n.Right.Extent.Text.Contains('Get-IdentityDeviceActivitySet')},$true)
$row=$ad;$nativeEntra=$null;$nativeManaged=$null
. ([scriptblock]::Create($deviceAssignment.Extent.Text))
Check ($null -eq (Merge-IdentityActivity $deviceActivity 90 -Now $now).Old) 'Missing device maps became not applicable'
Check (@($deviceActivity | Where-Object Status -eq 'Not applicable').Count -eq 0) 'Absent enrollment is not exclusion proof'

# Load actual accumulation, eligibility, and output projection functions only.
foreach($name in @('LoadCsv','AddLatest','Eligibility','V','SetWu')){
    $node=$pc.Find({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)
    . ([scriptblock]::Create($node.Extent.Text))
}
function Get-SourceCandidatePaths($Name){ @() }
function WriteLog { param($Message) }
$projectionScratch=Join-Path ([IO.Path]::GetTempPath()) ('identity-wu-projection-'+[guid]::NewGuid().ToString('N')+'.csv')
function Wu($State,$Policy,$Iso='',$Raw='2026-10-01T09:00:00Z') {
    [pscustomobject]@{PolicyId=$Policy;UpgradeEligibility=$State;AggregateState=$State;ExportDateTime=$Raw;ExportDateTimeIso8601=$Iso}
}
$older=Wu 'Capable' 'p1' '' '2026-10-01T09:00:00Z'
$newer=Wu 'Not Capable' 'p1' '' '2026-10-21T09:00:00Z'
$qualifiedOld=Wu 'Capable' 'p1' '2026-10-01T09:00:00Z'
$qualifiedNew=Wu 'Not Capable' 'p1' '2026-10-21T09:00:00Z'
$cases=@(
    @{Rows=@($older,$newer);Ambiguous=$true},
    @{Rows=@($qualifiedOld,$newer);Ambiguous=$true},
    @{Rows=@($qualifiedOld,$qualifiedNew);Ambiguous=$false},
    @{Rows=@($qualifiedOld,(Wu 'Not Capable' 'p1' '2026-10-01T09:00:00Z'));Ambiguous=$true},
    @{Rows=@($older,$older);Ambiguous=$false},
    @{Rows=@($older,(Wu 'Capable' 'p2'));Ambiguous=$true},
    @{Rows=@($older);Ambiguous=$false},
    @{Rows=@($qualifiedOld,$qualifiedNew,$qualifiedNew);Ambiguous=$false},
    @{Rows=@((Wu 'Capable' 'p1' '2099-10-21T09:00:00Z'));Ambiguous=$true}
)
$eligibilityAssignments=$pc.FindAll({param($n)$n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -in @('$upgradeRaw','$upgradeLabel','$w11EligibilityValue')},$true)
$guard=$pc.Find({param($n)$n -is [Management.Automation.Language.IfStatementAst] -and $n.Clauses[0].Item1.Extent.Text.Contains('$wuAnySelection.Status')},$true)
$sourceAssignment=$pc.Find({param($n)$n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$out[''Windows11UpgradeEligibilitySource'']'},$true)
try {
foreach($case in $cases){
    $representationResults=@()
    foreach($representation in @('PSCustomObject','ProjectedHashtable','OrderedDictionary')) {
    $results=@()
    foreach($reverse in @($false,$true)){
        $map=@{};$rows=@($case.Rows);if($reverse){[array]::Reverse($rows)}
        if($representation -eq 'ProjectedHashtable') {
            $rows | Export-Csv -LiteralPath $projectionScratch -NoTypeInformation
            $rows=@(LoadCsv 'synthetic-only.csv' @('PolicyId','UpgradeEligibility','AggregateState','ExportDateTime','ExportDateTimeIso8601') $projectionScratch)
            Check ($rows[0] -is [hashtable]) 'Real LoadCsv projection not exercised'
        } elseif($representation -eq 'OrderedDictionary') {
            $rows=@($rows | ForEach-Object {
                $record=[ordered]@{};foreach($property in $_.PSObject.Properties){$record[$property.Name]=$property.Value};$record
            })
        }
        foreach($item in $rows){AddLatest $map 'device' $item}
        $wuAnySelection=Resolve-IdentityWindowsUpdateRows $map.device
        $wuAny=$wuAnySelection.Row;$intune=$null;$isWin11=$false
        Check (($null -eq $wuAny) -eq $case.Ambiguous) 'Unexpected WU ambiguity'
        foreach($assignment in $eligibilityAssignments){. ([scriptblock]::Create($assignment.Extent.Text))}
        $migrationPhase='PHASE 1';$migrationReason='Synthetic independent inputs';$migrationAction='Synthetic action'
        . ([scriptblock]::Create($guard.Extent.Text))
        $out=[ordered]@{};. ([scriptblock]::Create($sourceAssignment.Extent.Text))
        SetWu $out $wuAny 'SyntheticPolicy'
        if($case.Ambiguous){Check ($w11EligibilityValue -eq 'Unknown' -and $migrationPhase -eq 'UNDETERMINED' -and $migrationAction -eq '6 - To qualify') 'Ambiguous WU still drives a recommendation'}
        $results+=([ordered]@{Eligibility=$w11EligibilityValue;Source=$out.Windows11UpgradeEligibilitySource;Phase=$migrationPhase;Action=$migrationAction;State=$out.WindowsUpdateSyntheticPolicyAggregateState;Status=$wuAnySelection.Status} | ConvertTo-Json -Compress)
    }
    Check ($results[0] -ceq $results[1]) 'Permutation changed a business decision'
    $representationResults+=$results[0]
    }
    Check (@($representationResults | Sort-Object -Unique).Count -eq 1) 'Row representation changed the decision'
}
} finally { Remove-Item -LiteralPath $projectionScratch -ErrorAction SilentlyContinue }
# The device-wide map must not collapse different configured policy grains.
$wuAnyByDevice=@{};$wuAnchorByDevice=@{};$wu24ByDevice=@{};$wu25ByDevice=@{}
$deviceKey='device';$anchorKey='p1';$policy24Key='p2';$policy25Key='p3'
$policyLoop=$pc.Find({param($n)$n -is [Management.Automation.Language.ForEachStatementAst] -and $n.Variable.Extent.Text -eq '$row' -and $n.Extent.Text.Contains('AddLatest $wuAnyByDevice')},$true)
$policyGuards=$policyLoop.Body.Statements | Where-Object { $_ -is [Management.Automation.Language.IfStatementAst] }
foreach($row in @($older,(Wu 'Not Capable' 'p2'))){
    $policyKey=$row.PolicyId;AddLatest $wuAnyByDevice $deviceKey $row
    foreach($statement in $policyGuards){. ([scriptblock]::Create($statement.Extent.Text))}
}
Check ($null -eq (Resolve-IdentityWindowsUpdateRows $wuAnyByDevice.device).Row) 'Different policies collapsed at device grain'
Check ((Resolve-IdentityWindowsUpdateRows $wuAnchorByDevice.device).Row.UpgradeEligibility -eq 'Capable') 'Anchor policy grain lost'
Check ((Resolve-IdentityWindowsUpdateRows $wu24ByDevice.device).Row.UpgradeEligibility -eq 'Not Capable') '24H2 policy grain lost'
Write-Output "PASS: $checks scope/Windows Update checks; actual consumer lists and AST rules; synthetic only."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAIhHvcCK4ZvvZy
# NEVQjl/EaIbGo/fwpmcYfIDx2H/qw6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIK/gyPXZ4lYDwQK2uLMcAMN9o16TILXIsVuCId7GY9ZMMA0GCSqG
# SIb3DQEBAQUABIIBgGU6DTTJBRxaObBn+7tnXgDjNAGomxh+aQ8zDImpWMKIzejM
# 8EwwBV+Pv8TTX5y+IY0B0Cx7gVV7sAFKvR8zgbKRtdZ7O7nZRd/JDs7bI3hMY05A
# LaPkzYkRi6nIeGQ5w9pJ9VT0oSoZZxMkOQsiYH/h/HkvontXskmnV097VmVc0LGi
# uQQ63hmd7Z1+OTiODJSCOUpgFE94TOhTcwaGnpg+1nAO9sKkVG1dpdMQCeqCB7Ti
# LdEEf2ai5GIvwAM/ySRk8EpP12QCCyW2M8Cbz3ua5MppkRwrlLo0Zmn9UFOhFjXk
# yioiiCcbmMsaF76IgKsmn4W0ZLWwy5F73Zz4U4/g+fED4M7ZsfCI17t2pSYj6k10
# P8FOMb7mZ/H+z1zlEChD0O97r/SlwmxXH4KpGhWYq6lwUO8rVpIHklJM07mFwECH
# Ley8JBOYbH6jG4UlQ9Biy2evLMmbNHaQ3bv/m1cMFVqECyIsRZxiadYWMhpmWey/
# tHFX3EJj/4PnlfMzo6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjcwOTQx
# NTFaMC8GCSqGSIb3DQEJBDEiBCDljxXv0pCyLmLFUGWVNqcE3/5OjZtCLzz7KfWM
# uZIsGTANBgkqhkiG9w0BAQEFAASCAgAmOUBi0qt4LH65UY68Ff6hAQpcRftNegxc
# cE7Jxcds4c7xJ/aGpc8yV0ZNerWkKjV6IPE3xYlfqg2g4k2RRNh80Y6JSrTIVfny
# Hm/JnSFuItHDto9FuKJB3qAPP/ju2hhlZzheQwjUUsrfspqrnNygH6fr+Irzemat
# KchKb3KsXYYxz2phX3NxEYcNSY64Ha4rde76c8VJ6Qi65RRV1SKntmv+9HgNDAJi
# Lalh4IvlspKsKtF1HeUzt1XCJ+Nyf098cr0WpyLhd8pvFAo5joA6xhjIfntsl1Wr
# qG57Aa293NvCaNISo2najpOAx4OrFixxiE0Vb6CRoGZKOUrXD5/RrMVxUu42wFbL
# KRs0TUsYU3Ns1ZpDp6Q9yEWQHWomNORVlg55AUE5n6OaBBHeMlCO8OIUIKGXJmKY
# V342O8W7OfHwKuFajEsl9oti+AAilKdJFJiAq7o7cQKHEUlubtW5/hAccsgNcfFs
# p1KXOsMPPdqt6970nofxN9xVZsVd9Fueush72+pHTPyqDwN2c3AhaIAQ/+aPzEQU
# qfjcPlKdURnPhhlKZb8Zw2hwWO78XEmjgnrTh4S+v/Gwy+GYAnzBpfKuVOUOkOKb
# jdp0Tu7e1cLP/OFZGc3ZnCWegBfnekfMLzyn3G0JCxIG4C4Ecc3JeZtCKfePK14k
# +nUImEhBHg==
# SIG # End signature block
