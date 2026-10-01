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
