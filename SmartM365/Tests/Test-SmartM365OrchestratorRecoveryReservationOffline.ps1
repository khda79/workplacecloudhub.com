#requires -Version 7.0
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot '../SmartInventory/Orchestrator/SmartM365.Orchestrator.Pipeline.psm1') -Force
$root=Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-RecoveryFence-' + [guid]::NewGuid().ToString('N'))
try {
    $null=New-Item -ItemType Directory -Path $root -Force
    $selection=[pscustomobject]@{SelectedJobs=@([pscustomobject]@{Name='Example'})}
    $request=New-SmartM365OrchestratorPipelineRequest -SharedDataFolderPath $root -Tenant test -Pipeline Jobs -Selection $selection -ManifestHash 'fixture' -Recovery
    if (-not (Test-SmartM365OrchestratorRecoveryReservation -SharedDataFolderPath $root -JobName Example -Tenant test)) { throw 'Pending recovery did not reserve its job.' }
    if (Test-SmartM365OrchestratorRecoveryReservation -SharedDataFolderPath $root -JobName Other -Tenant test) { throw 'An unrelated job was reserved.' }
    $statusPath=Get-SmartM365OrchestratorPipelineJobStatusPath -SharedDataFolderPath $root -BatchId $request.BatchId -JobName Example
    $saved=[IO.File]::ReadAllBytes($statusPath)
    Remove-Item -LiteralPath $statusPath -Force
    if (-not (Test-SmartM365OrchestratorRecoveryReservation -SharedDataFolderPath $root -JobName Example -Tenant test)) { throw 'Missing status must keep the reservation closed.' }
    $null=New-Item -ItemType Directory -Path (Join-Path $root 'Config'),(Join-Path $root 'SERVER-A') -Force
    [IO.File]::WriteAllText((Join-Path $root 'Config/Orchestrator-Cluster.json.txt'),'{"ExpectedOrchestratorServers":["SERVER-A"]}')
    [IO.File]::WriteAllText((Join-Path $root 'SERVER-A/Orchestrator-Heartbeat.json.txt'),'{"PipelineQueueProtocol":1}')
    $overlapSelection=[pscustomobject]@{SelectedJobs=@([pscustomobject]@{Name='Example'},[pscustomobject]@{Name='Other'})}
    $refused=$false; $refusal=''
    try { $null=New-SmartM365OrchestratorPipelineRequest -SharedDataFolderPath $root -Tenant test -Pipeline Jobs -Selection $overlapSelection -ManifestHash 'fixture' } catch { $refusal=$_.Exception.Message; $refused=$refusal -match 'overlaps' }
    if (-not $refused) { throw "Missing status must prevent a differently selected overlapping request: $refusal" }
    [IO.File]::WriteAllBytes($statusPath,$saved)
    $null=Set-SmartM365OrchestratorPipelineJobStatus -SharedDataFolderPath $root -BatchId $request.BatchId -JobName Example -Status Running
    $requestDocument=Get-Content -LiteralPath $request.RequestPath -Raw | ConvertFrom-Json
    $requestDocument | Add-Member -NotePropertyName Cancellation -NotePropertyValue ([pscustomobject]@{Reason='synthetic'})
    [IO.File]::WriteAllText($request.RequestPath,($requestDocument | ConvertTo-Json -Depth 20))
    $refused=$false
    try { $null=New-SmartM365OrchestratorPipelineRequest -SharedDataFolderPath $root -Tenant test -Pipeline Jobs -Selection $overlapSelection -ManifestHash 'fixture' } catch { $refused=$_.Exception.Message -match 'overlaps' }
    if (-not $refused) { throw 'A cancelling recovery must still protect its running job.' }
    $null=Set-SmartM365OrchestratorPipelineJobStatus -SharedDataFolderPath $root -BatchId $request.BatchId -JobName Example -Status Success
    if (Test-SmartM365OrchestratorRecoveryReservation -SharedDataFolderPath $root -JobName Example -Tenant test) { throw 'A completed recovery still holds the reservation.' }
    'PASS: recovery reservation and missing-status protection'
}
finally { if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force } }
