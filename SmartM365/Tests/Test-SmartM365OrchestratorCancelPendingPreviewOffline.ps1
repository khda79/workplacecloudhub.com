#requires -Version 7.0
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot '../SmartInventory/Orchestrator/SmartM365.Orchestrator.Insights.psm1') -Force
$operations = [pscustomobject]@{
    Servers=@([pscustomobject]@{Mode='Active';State='Online';StatePersistence='OK'})
    Maintenance=[pscustomobject]@{Enabled=$false;Revision=7;CatchUpWindows=@([pscustomobject]@{Revision=5})}
    MaintenanceError=''; Running=@([pscustomobject]@{Job='RunningJob'}); Pending=@([pscustomobject]@{Job='PendingJob'})
}
$requests = @([pscustomobject]@{Status='Running';BatchId='one'},[pscustomobject]@{Status='Queued';BatchId='two'},[pscustomobject]@{Status='Success';BatchId='three'})
$preview = Get-SmartM365OrchestratorCancelPendingPreview -Operations $operations -Requests $requests
if (@($preview.Running).Count -ne 1 -or @($preview.Pending).Count -ne 1 -or @($preview.ActiveRequests).Count -ne 2 -or @($preview.CatchUpWindows).Count -ne 1 -or $preview.MaintenanceRevision -ne 7) { throw 'Cancellation preview omitted work or running jobs' }
$operations.Servers[0].State='Stale'
$refused=$false
try { $null=Get-SmartM365OrchestratorCancelPendingPreview -Operations $operations -Requests $requests } catch { $refused=$true }
if (-not $refused) { throw 'Stale active server must not be presented as cancellable' }
$operations.Servers[0].State='Online'; $operations.Maintenance.Enabled=$true
$refused=$false
try { $null=Get-SmartM365OrchestratorCancelPendingPreview -Operations $operations -Requests $requests } catch { $refused=$true }
if (-not $refused) { throw 'An existing maintenance session must not be overwritten' }
'PASS: read-only cancellation preview'
