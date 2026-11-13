#requires -Version 7.0
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot '../SmartInventory/Orchestrator/SmartM365.Orchestrator.Insights.psm1') -Force
$jobs = [pscustomobject]@{ Jobs = @(
    [pscustomobject]@{Name='EXO';Group='Exchange';Enabled=$true;AssignmentMode='Elected';DependsOn=@();RequiredCapabilities=@('EXO')},
    [pscustomobject]@{Name='Licences';Group='M365';Enabled=$true;AssignmentMode='Elected';DependsOn=@('EXO');RequiredCapabilities=@('Graph')},
    [pscustomobject]@{Name='Prepare';Group='M365';Enabled=$true;AssignmentMode='Elected';DependsOn=@('Licences');RequiredCapabilities=@('Python')}
) }
$operations = [pscustomobject]@{ Servers=@([pscustomobject]@{Mode='Active';State='Online'}); Pending=@([pscustomobject]@{Job='Prepare';ServerState='Online'},[pscustomobject]@{Job='Licences';ServerState='Online'}); Running=@([pscustomobject]@{Job='EXO'}) }
$plan = [pscustomobject]@{ Assignments=@([pscustomobject]@{JobName='Licences';OwnerServer='A'},[pscustomobject]@{JobName='Prepare';OwnerServer='B'}) }
$preview = Get-SmartM365OrchestratorRecoveryPreview -JobsDocument $jobs -Operations $operations -Plan $plan -History @() -AllBlocked
if (@($preview.RequestedJobs).Count -ne 2 -or @($preview.Jobs).Count -ne 3) { throw 'Blocked selection did not include unique requested jobs and dependencies' }
$exo = @($preview.Jobs | Where-Object Name -eq 'EXO')[0]
if (-not $exo.Running -or -not $exo.OwnerUnavailable -or -not $exo.IncludedAsDependency) { throw 'Preview did not report running and unavailable dependency' }
$selected = Get-SmartM365OrchestratorRecoveryPreview -JobsDocument $jobs -Operations $operations -Plan $plan -JobName @('Licences')
if (@($selected.Jobs).Count -ne 2 -or @($selected.Jobs | Where-Object Name -eq 'Prepare').Count) { throw 'Selected recovery included an unrelated job' }
$operations.Pending=@()
$history = @([pscustomobject]@{JobName='Prepare';Status='BlockedDependencyTimeout'},[pscustomobject]@{JobName='Prepare';Status='Success'},[pscustomobject]@{JobName='Licences';Status='Success'},[pscustomobject]@{JobName='Licences';Status='BlockedDependencyFailed'})
$terminal = Get-SmartM365OrchestratorRecoveryPreview -JobsDocument $jobs -Operations $operations -Plan $plan -History $history -AllBlocked
if ((@($terminal.RequestedJobs) -join ',') -ne 'Prepare') { throw 'Latest terminal dependency block was not selected correctly' }
$operations.Pending=@([pscustomobject]@{Job='Prepare';ServerState='Online'})
$operations.Pending[0].ServerState='Stale'
$refused=$false
try { $null = Get-SmartM365OrchestratorRecoveryPreview -JobsDocument $jobs -Operations $operations -Plan $plan -History @() -AllBlocked } catch { $refused=$true }
if (-not $refused) { throw 'Stale pending state must not be presented as a complete selection' }
'PASS: read-only recovery preview'
