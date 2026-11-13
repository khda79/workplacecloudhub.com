Import-Module (Join-Path $PSScriptRoot '../SmartInventory/Orchestrator/SmartM365.Orchestrator.Distributed.psm1') -Force
$job = [pscustomobject]@{ Name='EXO-Fast'; Enabled=$true; AssignmentMode='Elected'; RequiredCapabilities=@('EXO','Graph'); RequiredGraphAppRoles=@(); EstimatedDurationMinutes=10; Schedule=[pscustomobject]@{Type='Daily';Times=@('00:15');DaysOfWeek=@()} }
$server = [pscustomobject]@{ ServerName='A'; ReadyCapabilities=@('Graph'); GraphAppRoles=@(); GeneratedAtUtc='2026-11-13T03:41:00Z'; Results=@([pscustomobject]@{Name='EXO';Ready=$false;Detail='Connection timed out'}) }
$plan=Get-SmartM365OrchestratorElectionPlan -Jobs @($job) -ServerCapabilities @($server)
if (@($plan.Assignments).Count -ne 0 -or @($plan.UnassignedGroups).Count -ne 1 -or $plan.UnassignedGroups[0].Reason -notmatch 'EXO unavailable.*Connection timed out') { throw 'Failure diagnostic or fail-closed election is wrong' }
$server.ReadyCapabilities=@('Graph','EXO')
$server.Results=@([pscustomobject]@{Name='EXO';Ready=$true;Detail='OK'})
$plan=Get-SmartM365OrchestratorElectionPlan -Jobs @($job) -ServerCapabilities @($server) -PreviousPlan $plan -PreservePreviousOwners
if (@($plan.Assignments).Count -ne 1 -or @($plan.UnassignedGroups).Count -ne 0) { throw 'Recovered capability did not restore assignment' }
'PASS: election diagnosis and recovery'
