<#
.SYNOPSIS
Offline tests for the orchestrator shared dependency gate (latest occurrence, overlap coverage) and mixed-capability election.
.VERSION
1.1
#>
#requires -Version 7.0
[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
$base=Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
Import-Module (Join-Path $base 'SmartM365.Orchestrator.Distributed.psm1') -Force
$transport=Get-Module SmartM365.JsonTransport
$policy=& $transport {(Get-Command Get-SmartM365JsonTransportPolicy).ScriptBlock}
$ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $base 'SmartM365-Inventory-Orchestrator.ps1'),[ref]$null,[ref]$null)
$defs=foreach($name in @('Get-JobOccurrencesInWindow','Get-LatestPastOccurrence','Get-OrchestratorSharedDependencyStatus','Test-OrchestratorOccurrenceCoveredByOverlap','Get-OrchestratorClaimOccurrenceUtc','ConvertTo-OrchestratorUtcTime')) {
    $ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true).Extent.Text
}
$gate=New-Module -ScriptBlock ([scriptblock]::Create($defs -join "`n"))
$root=Join-Path $env:TEMP ('shared-deps-'+[guid]::NewGuid().ToString('N'))
$claims=Join-Path $root 'Claims';$leases=Join-Path $root 'Leases'
New-Item -ItemType Directory $claims,$leases|Out-Null
$script:passed=0
function Check($Value,$Message){if(-not $Value){throw $Message};$script:passed++}
try {
    & $transport {function script:Get-SmartM365JsonTransportPolicy {@{Mode='JsonText';QualifiedUncRoots=@()}}}
    & $gate {param($c,$l) $script:Settings=[pscustomobject]@{ElectionClaimsPath=$c;ConcurrencyLeasesPath=$l;ElectionClaimGraceMinutes=15};$script:DependencyFreshCache=@{}} $claims $leases
    $now=[datetime]::new(2026,9,28,10,0,0,[DateTimeKind]::Local)
    $job=[pscustomobject]@{Name='OnPremParent';ConcurrencyKey='OnPremParent';Schedule=[pscustomobject]@{Type='Daily';Times=@('03:00');DaysOfWeek=@()}}
    function Status { & $gate {param($j,$n) Get-OrchestratorSharedDependencyStatus -Job $j -Now $n} $job $now }
    Check ((Status) -eq 'Waiting') 'Absent dependency allowed consumer.'
    $yesterday=Enter-SmartM365OrchestratorOccurrenceClaim -ClaimsRootPath $claims -JobName $job.Name -Occurrence $now.Date.AddDays(-1).AddHours(3) -OwnerServer REMOTE -PlanId synthetic
    Set-SmartM365OrchestratorOccurrenceClaim -ClaimPath $yesterday.ClaimPath -OwnerServer REMOTE -Status Success|Out-Null
    Check ((Status) -eq 'Waiting') 'Older success substituted for latest scheduled occurrence.'
    $current=Enter-SmartM365OrchestratorOccurrenceClaim -ClaimsRootPath $claims -JobName $job.Name -Occurrence $now.Date.AddHours(3) -OwnerServer REMOTE -PlanId synthetic
    foreach($state in @('Claimed','Running','RetryScheduled')) {
        Set-SmartM365OrchestratorOccurrenceClaim -ClaimPath $current.ClaimPath -OwnerServer REMOTE -Status $state|Out-Null
        Check ((Status) -eq 'Waiting') "Remote $state dependency allowed consumer."
    }
    foreach($state in @('Failed','TimedOut','Interrupted')) {
        Set-SmartM365OrchestratorOccurrenceClaim -ClaimPath $current.ClaimPath -OwnerServer REMOTE -Status $state|Out-Null
        Check ((Status) -eq 'Failed') "Remote $state dependency treated as success."
    }
    foreach($state in @('Success','CompletedWithWarnings')) {
        Set-SmartM365OrchestratorOccurrenceClaim -ClaimPath $current.ClaimPath -OwnerServer REMOTE -Status $state|Out-Null
        Check ((Status) -eq 'Ready') "Remote $state not accepted."
    }
    $lease=Enter-SmartM365OrchestratorConcurrencyLease -LeasesRootPath $leases -ConcurrencyKey $job.ConcurrencyKey -JobName $job.Name -Occurrence $now -OwnerServer REMOTE
    Check ((Status) -eq 'Waiting') 'New remote refresh ignored while older scheduled claim succeeded.'
    Exit-SmartM365OrchestratorConcurrencyLease -LeasePath $lease.LeasePath -LeaseId $lease.Lease.LeaseId -OwnerServer REMOTE|Out-Null
    Check ((Status) -eq 'Ready') 'Released remote refresh still blocked consumer.'
    $bytes=[IO.File]::ReadAllBytes($current.ClaimPath)
    [IO.File]::WriteAllText($current.ClaimPath,'{')
    $rejected=$false;try{Status|Out-Null}catch{$rejected=$true}
    Check $rejected 'Invalid preferred dependency claim accepted.'
    [IO.File]::WriteAllBytes($current.ClaimPath,$bytes)
    $job.Schedule=[pscustomobject]@{Type='Weekly';Times=@('03:00');DaysOfWeek=@('Sunday')}
    Check ((Status) -eq 'Ready') 'Weekly dependency did not select latest weekly occurrence.'
    $job.Schedule=[pscustomobject]@{Type='Manual';Times=@();DaysOfWeek=@()}
    Check ((Status) -eq 'Waiting') 'Unscheduled dependency silently allowed consumer.'

    # Overlap guard: the 00:00 occurrence was skipped (no claim) while the 17:00 run was in progress.
    $overlap=[pscustomobject]@{Name='OverlapParent';ConcurrencyKey='';TimeoutMinutes=170;Schedule=[pscustomobject]@{Type='Daily';Times=@('00:00','17:00');DaysOfWeek=@()}}
    $overlapNow=[datetime]::new(2026,10,1,1,30,0,[DateTimeKind]::Local)
    $previousOccurrence=[datetime]::new(2026,9,30,17,0,0,[DateTimeKind]::Local)
    function OverlapStatus { & $gate {param($j,$n) $script:DependencyFreshCache=@{}; Get-OrchestratorSharedDependencyStatus -Job $j -Now $n} $overlap $overlapNow }
    function Write-OverlapClaim([string]$Status,[datetime]$CreatedLocal,[datetime]$UpdatedLocal) {
        $folder=Join-Path $claims 'OverlapParent'; New-Item -ItemType Directory $folder -Force|Out-Null
        Get-ChildItem $folder -File|Remove-Item -Force
        $claim=[ordered]@{SchemaVersion=1;ClaimId=[guid]::NewGuid().ToString('N');JobName='OverlapParent';OccurrenceUtc=$previousOccurrence.ToUniversalTime().ToString('o');OwnerServer='REMOTE';PlanId='synthetic';OrchestratorPid=1;Status=$Status;Attempt=0;CreatedAtUtc=$CreatedLocal.ToUniversalTime().ToString('o');UpdatedAtUtc=$UpdatedLocal.ToUniversalTime().ToString('o');SafeUntilUtc=$UpdatedLocal.ToUniversalTime().AddHours(1).ToString('o')}
        $path=Join-Path $folder ($previousOccurrence.ToUniversalTime().ToString('yyyyMMddTHHmmssfffZ')+'.json.txt')
        Set-Content -LiteralPath $path -Value ($claim|ConvertTo-Json) -Encoding utf8
    }
    $late=[datetime]::new(2026,9,30,23,33,0,[DateTimeKind]::Local); $afterMidnight=[datetime]::new(2026,10,1,0,16,0,[DateTimeKind]::Local)
    Write-OverlapClaim 'Success' $late $afterMidnight
    Check ((OverlapStatus) -eq 'Ready') 'A successful run in progress at the skipped occurrence did not cover it.'
    Write-OverlapClaim 'Success' $previousOccurrence $previousOccurrence.AddMinutes(45)
    Check ((OverlapStatus) -eq 'Waiting') 'A run that ended before the occurrence covered it.'
    Write-OverlapClaim 'Failed' $late $afterMidnight
    Check ((OverlapStatus) -eq 'Waiting') 'A failed overlapping run covered the occurrence.'
    Write-OverlapClaim 'Running' $late $afterMidnight
    Check ((OverlapStatus) -eq 'Waiting') 'A still running overlapping run covered the occurrence.'
    $currentOccurrence=[datetime]::new(2026,10,1,0,0,0,[DateTimeKind]::Local)
    Write-OverlapClaim 'Success' $late $afterMidnight
    $own=Enter-SmartM365OrchestratorOccurrenceClaim -ClaimsRootPath $claims -JobName 'OverlapParent' -Occurrence $currentOccurrence -OwnerServer REMOTE -PlanId synthetic
    Set-SmartM365OrchestratorOccurrenceClaim -ClaimPath $own.ClaimPath -OwnerServer REMOTE -Status Failed|Out-Null
    Check ((OverlapStatus) -eq 'Failed') 'An existing claim for the latest occurrence was overridden by an older run.'

    $jobs=@(
        [pscustomobject]@{Name='OnPremParent';Enabled=$true;AssignmentMode='Elected';DependsOn=@();RequiredCapabilities=@('SharedRuntime','ExchangeOnPrem');RequiredGraphAppRoles=@();EstimatedDurationMinutes=10;Schedule=[pscustomobject]@{Type='Daily';Times=@('03:00');DaysOfWeek=@()}},
        [pscustomobject]@{Name='CloudChild';Enabled=$true;AssignmentMode='Elected';DependsOn=@('OnPremParent');RequiredCapabilities=@('SharedRuntime','Graph','TeamsPowerShell');RequiredGraphAppRoles=@('Reports.Read.All');EstimatedDurationMinutes=10;Schedule=[pscustomobject]@{Type='Daily';Times=@('03:10');DaysOfWeek=@()}}
    )
    $servers=@(
        [pscustomobject]@{ServerName='EXCHANGE';ReadyCapabilities=@('SharedRuntime','ExchangeOnPrem');GraphAppRoles=@()},
        [pscustomobject]@{ServerName='CLOUD';ReadyCapabilities=@('SharedRuntime','Graph','TeamsPowerShell');GraphAppRoles=@('Reports.Read.All')}
    )
    $plan=Get-SmartM365OrchestratorElectionPlan -Jobs $jobs -ServerCapabilities $servers -ServerJobPolicies @{EXCHANGE=@{OnlyJobsRequiring=@('ExchangeOnPrem')}}
    Check ($plan.UnassignedGroups.Count -eq 0 -and $plan.Assignments.Count -eq 2) 'Mixed dependency group remained unassigned.'
    Check (($plan.Assignments|Where-Object JobName -eq 'OnPremParent').OwnerServer -eq 'EXCHANGE') 'On-prem parent assigned to cloud host.'
    Check (($plan.Assignments|Where-Object JobName -eq 'CloudChild').OwnerServer -eq 'CLOUD') 'Cloud child bypassed Exchange server policy.'
    $servers[1].GraphAppRoles=@()
    $restricted=Get-SmartM365OrchestratorElectionPlan -Jobs $jobs -ServerCapabilities $servers -ServerJobPolicies @{EXCHANGE=@{OnlyJobsRequiring=@('ExchangeOnPrem')}}
    Check ($restricted.Assignments.Count -eq 1 -and $restricted.UnassignedGroups.Count -eq 1) 'Missing Graph role bypassed.'
    [pscustomobject]@{Passed=$script:passed;Fixture=$root;Scope='Synthetic files and topology; no production jobs or network access'}
} finally {
    & $transport {param($p) Set-Item Function:script:Get-SmartM365JsonTransportPolicy $p} $policy
    Remove-Module $gate -Force
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCIxJUd+OtYnWzu
# 1MWj0nK1ZhVRUZBPRtBD1C9j/tkCqKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIG6jfwLrxJpQb/BkY5WdPjuwIDhhn8z2n0OeiVdCLxcqMA0GCSqG
# SIb3DQEBAQUABIIBgA+jvvkq/VwwIhwarMn3iOPVEZqR1tcboExrQ2n6QNtOgI/N
# /UCmxqiMFxoQuHdR0rNy/Bs6cuLH41cN4KaUT9FAMcGPmiB779kkKIJ2myzxaW7Q
# 27PLz8OPW/WVZB8ShDEWYPrQQS667w3Lq/6mJWSsjKvKTvI/LmHrHKx9LADZN3tx
# 02EtmoR4a4/RJvfBHLYP/n5paPsYoRlJfo7R7BsmFSwG4EcdBxe1J2uVRKMqKCxt
# t41NHjHdDuMvavGSDXXkt1qH6ENjczuxakNB8jMfJ6pjf0lRPY1tsjxc1PBn3aHI
# 81noYjAVb5c3L3TM9Xxr/+WB8x0biydTIXKctcchtKFoHDpVUA+4DGNLM2UTICZA
# pTfrd68qFwI5ODQBrtwKa1wBF4bVEOpbisXpKG1pwY9+SfESp4bW2mVVtIG91h50
# jmlyx17kEEftSmzLzNRMv4PmPUK4BDIJlRgz/gXlQBWe1/KVCNkVr7koO2J2LJZ1
# BIHe3Z0cHf/ajUrfaKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDExMDUx
# MDlaMC8GCSqGSIb3DQEJBDEiBCD9wS66E5iBlm/mzRBzxM8IwuaazNrNFwTTbj9y
# GOYMdDANBgkqhkiG9w0BAQEFAASCAgB3GgvMD/qRQdLphVTipWT0Dm0yBDyZpSMO
# YAR3EdOOC07uglKbsZB5iUMHk5fnmWaP9nDW3NpOjG81ggK45lpU0D+dUyjiUAio
# jMdu7VYbgKNtQuNZSibAWE/yDs8Ns2HDCaF+EmI+L5nzi396uXo6Sg3/0Cequqnx
# IH49z5snGT9ORChBVzL3OaH2u61i/U8j2YObtvvNDa1P7t2GRd4dcSqiJl711LsV
# orR3GH+x2zKm5oRugUk+O0OZWcxicTh0wLQzgXg1KA2YQTlCQ8zgz1pwZW8FZhPo
# dSkGL5UbboKVIbyPfl1U7l53TeTDqSSaMsxBKQLuKW96ZrQOhAFfZPneakooscyo
# 9M2rGOy4NJ8H7QQhzOLObZHffuLLjj9q4S8uWFFQWznrC2WF9Lxl9EF2OCQGCQyZ
# r8iznt3wF5UWPF+tpoTWZmgnPpRnRQVM7ssdtcwgpxGiB3PoRwndJeObxUWP9w1s
# 91cCZMCKWeHo8W40Wt22pTddxLx9/mC+DREuURyQichzc1otsgF2PUEh9tgPUZ26
# 3ZgV/DncEpJFXHzN7eFQ9CxdM0O3adDRPFDSlMMjTzXMQcR3NKOBiRIirTTBEk8E
# tGDKxtXzleFp7dsV9fFnUGMbUKWcpcr5MzMM8Rq2TURolhZuPj2SS2hezvcHzuDL
# 2us06D4Syg==
# SIG # End signature block
