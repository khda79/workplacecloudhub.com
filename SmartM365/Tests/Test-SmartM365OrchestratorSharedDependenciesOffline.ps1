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
Import-Module (Join-Path $base 'SmartM365.Orchestrator.Insights.psm1') -Force
$transport=Get-Module SmartM365.JsonTransport
$policy=& $transport {(Get-Command Get-SmartM365JsonTransportPolicy).ScriptBlock}
$ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $base 'SmartM365-Inventory-Orchestrator.ps1'),[ref]$null,[ref]$null)
$defs=foreach($name in @('Get-JobOccurrencesInWindow','Get-LatestPastOccurrence','Get-OrchestratorSharedDependencyStatus','Get-OrchestratorLaterDependencyClaimStatus','Test-OrchestratorOccurrenceCoveredByOverlap','Get-OrchestratorClaimOccurrenceUtc','ConvertTo-OrchestratorUtcTime')) {
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

    # A manual pipeline run after the scheduled occurrence satisfies that day's dependency.
    $manual=[pscustomobject]@{Name='ManualRecoveryParent';ConcurrencyKey='ManualRecoveryParent';Schedule=[pscustomobject]@{Type='Daily';Times=@('00:10');DaysOfWeek=@()}}
    $manualNow=[datetime]::new(2026,10,8,16,0,0,[DateTimeKind]::Local)
    $scheduled=$manualNow.Date.AddMinutes(10)
    # Pipeline claim JSON keeps 100 ns, while the filename keeps only milliseconds.
    $recovery=$manualNow.Date.AddHours(2).AddMinutes(58).AddSeconds(48).AddMilliseconds(320).AddTicks(7090)
    $manualFolder=Join-Path $claims $manual.Name
    New-Item -ItemType Directory $manualFolder -Force|Out-Null
    function ManualStatus { & $gate {param($j,$n) Get-OrchestratorSharedDependencyStatus -Job $j -Now $n} $manual $manualNow }
    function Write-ManualClaim([datetime]$Occurrence,[string]$Status,[datetime]$Created,[datetime]$Updated) {
        $document=[ordered]@{SchemaVersion=1;ClaimId=[guid]::NewGuid().ToString('N');JobName=$manual.Name;OccurrenceUtc=$Occurrence.ToUniversalTime().ToString('o');OwnerServer='REMOTE';PlanId='synthetic';OrchestratorPid=1;Status=$Status;Attempt=0;CreatedAtUtc=$Created.ToUniversalTime().ToString('o');UpdatedAtUtc=$Updated.ToUniversalTime().ToString('o')}
        $path=Join-Path $manualFolder ($Occurrence.ToUniversalTime().ToString('yyyyMMddTHHmmssfffZ')+'.json.txt')
        [IO.File]::WriteAllText($path,($document|ConvertTo-Json),[Text.UTF8Encoding]::new($false))
    }
    Check ((ManualStatus) -eq 'Waiting') 'Missing scheduled and recovery claims allowed consumer.'
    Write-ManualClaim $recovery 'Running' $recovery $recovery.AddMinutes(1)
    Check ((ManualStatus) -eq 'Waiting') 'In-progress recovery allowed consumer.'
    Write-ManualClaim $recovery 'Failed' $recovery $recovery.AddMinutes(2)
    Check ((ManualStatus) -eq 'Failed') 'Failed recovery allowed consumer.'
    Write-ManualClaim $recovery 'Success' $recovery $recovery.AddHours(2)
    Check ((ManualStatus) -eq 'Ready') 'Successful later recovery did not satisfy the scheduled dependency.'
    $manualLease=Enter-SmartM365OrchestratorConcurrencyLease -LeasesRootPath $leases -ConcurrencyKey $manual.ConcurrencyKey -JobName $manual.Name -Occurrence $recovery -OwnerServer REMOTE
    Check ((ManualStatus) -eq 'Waiting') 'Active recovery lease allowed consumer.'
    Exit-SmartM365OrchestratorConcurrencyLease -LeasePath $manualLease.LeasePath -LeaseId $manualLease.Lease.LeaseId -OwnerServer REMOTE|Out-Null
    Write-ManualClaim $scheduled 'Failed' $scheduled $scheduled.AddMinutes(1)
    Check ((ManualStatus) -eq 'Ready') 'Later successful recovery did not override the failed scheduled claim.'
    Write-ManualClaim $scheduled 'Running' $scheduled $scheduled.AddMinutes(1)
    Check ((ManualStatus) -eq 'Waiting') 'Later recovery bypassed a still-running scheduled occurrence.'
    Write-ManualClaim $scheduled 'Success' $scheduled $scheduled.AddMinutes(1)
    Write-ManualClaim $recovery 'Failed' $recovery $recovery.AddMinutes(2)
    Check ((ManualStatus) -eq 'Ready') 'Later failed manual run invalidated a successful scheduled claim.'
    $insightsRoot=Join-Path $root 'Shared'
    $insightsFolder=Join-Path $insightsRoot 'Election/Claims/ManualRecoveryParent'
    New-Item -ItemType Directory $insightsFolder -Force|Out-Null
    Copy-Item -LiteralPath (Join-Path $manualFolder ($scheduled.ToUniversalTime().ToString('yyyyMMddTHHmmssfffZ')+'.json.txt')) -Destination $insightsFolder
    Copy-Item -LiteralPath (Join-Path $manualFolder ($recovery.ToUniversalTime().ToString('yyyyMMddTHHmmssfffZ')+'.json.txt')) -Destination $insightsFolder
    $consumer=[pscustomobject]@{Name='ManualRecoveryConsumer';DependsOn=@($manual.Name);DependencyMode='LatestOccurrence'}
    $manual|Add-Member -NotePropertyName Enabled -NotePropertyValue $true
    $manual|Add-Member -NotePropertyName AssignmentMode -NotePropertyValue Elected
    $jobsDocument=[pscustomobject]@{Jobs=@($manual,$consumer)}
    $readiness=@(Get-SmartM365OrchestratorDependencyReadiness -SharedDataFolderPath $insightsRoot -JobsDocument $jobsDocument -JobName $consumer.Name -Now $manualNow)
    Check ($readiness.Count -eq 1 -and $readiness[0].State -eq 'Ready') 'GUI readiness disagrees with a successful scheduled claim.'
    Write-ManualClaim $scheduled 'Failed' $scheduled $scheduled.AddMinutes(1)
    Copy-Item -LiteralPath (Join-Path $manualFolder ($scheduled.ToUniversalTime().ToString('yyyyMMddTHHmmssfffZ')+'.json.txt')) -Destination $insightsFolder -Force
    Write-ManualClaim $recovery 'Success' $recovery $recovery.AddHours(2)
    Copy-Item -LiteralPath (Join-Path $manualFolder ($recovery.ToUniversalTime().ToString('yyyyMMddTHHmmssfffZ')+'.json.txt')) -Destination $insightsFolder -Force
    $readiness=@(Get-SmartM365OrchestratorDependencyReadiness -SharedDataFolderPath $insightsRoot -JobsDocument $jobsDocument -JobName $consumer.Name -Now $manualNow)
    Check ($readiness.Count -eq 1 -and $readiness[0].State -eq 'Ready' -and $readiness[0].Detail -match 'later run') 'GUI readiness did not acknowledge successful recovery.'
    Write-ManualClaim $scheduled 'Running' $scheduled $scheduled.AddMinutes(1)
    Copy-Item -LiteralPath (Join-Path $manualFolder ($scheduled.ToUniversalTime().ToString('yyyyMMddTHHmmssfffZ')+'.json.txt')) -Destination $insightsFolder -Force
    $readiness=@(Get-SmartM365OrchestratorDependencyReadiness -SharedDataFolderPath $insightsRoot -JobsDocument $jobsDocument -JobName $consumer.Name -Now $manualNow)
    Check ($readiness.Count -eq 1 -and $readiness[0].State -eq 'Waiting') 'GUI readiness bypassed a running scheduled occurrence.'

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
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBWbtbjQsYIWtp/
# mcspBdWCGyygRDAKxLk+0XxS/u5Y3KCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDAPXZkfwDe+AsfJdfKjIhr
# 0Zt6BvfqybLgwT8crYaJMDANBgkqhkiG9w0BAQEFAASCAYA+nxBxfNeMz8FWKbl4
# CLIm/icw0xcIyWYMYG2c2DtiJv3teMCh8w0vfxUIVQVVi4izCGPX4qQ3jiIlBgAH
# 4v6o4ybLY17tMiiDbeolfuP8yksvdTIa0OzbN8aB8VHvxKKVRzeQ7QMir6h3ebMA
# 5hC3kDxTbvryX7sBAwNNlqat4woS+Us7CO4AfX+IfJ0Gd+pKHboqOU2BbyMfJVPT
# 4YciS0uMCxLQCd0HKOWoIRj/bArNa/2gr3fm10XNVtd/w/W/6aQx2nzgQlDpwiYi
# y4Y5nz9YpNGpvi8ceHTXyxtFBvNEvLu1JriAZGQ1yAvCRWAvT/NRceCLx7fcox9c
# X6SiL3JcQJ756/xReXMVF86fpAoZdJSDcx4Zy+hiM2JeKhZ6JBna8jR79Wt6s0jo
# DRcImRu9iHxi9mIRgcKHt5wHeuvsRXOlN03sMs/xAXkMBn+08v1OtO3am7AygdDD
# 7NTKFDjfQWbPxNyIGjRjAvf4EoPfgjxbU5IMncr3C3m/gYM=
# SIG # End signature block
