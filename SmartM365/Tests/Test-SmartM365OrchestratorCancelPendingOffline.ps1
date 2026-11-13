#requires -Version 7.0
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot '../SmartInventory/Orchestrator/SmartM365.Orchestrator.PendingWork.psm1') -Force
Import-Module (Join-Path $PSScriptRoot '../SmartInventory/Orchestrator/SmartM365.Orchestrator.Maintenance.psm1')
Import-Module (Join-Path $PSScriptRoot '../SmartInventory/Orchestrator/SmartM365.Orchestrator.Pipeline.psm1')
$root=Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-CancelPending-' + [guid]::NewGuid().ToString('N'))
$worker=$null
function Write-Fixture($Path,$Value) {
    $null=New-Item -ItemType Directory -Path (Split-Path $Path -Parent) -Force
    [IO.File]::WriteAllText($Path,($Value | ConvertTo-Json -Depth 15),[Text.UTF8Encoding]::new($false))
}
function New-Heartbeat($control) {
    [pscustomobject]@{ Tenant='test';ScriptVersion='1.5.90';Timestamp=[datetime]::UtcNow.ToString('o');Lifecycle='Running';RunningJobs=@([pscustomobject]@{Name='AlreadyRunning';Pid=999})
        MaintenanceProtocol=1;MaintenanceCatchUpProtocol=1;MaintenanceResetProtocol=1;PipelineCancellationProtocol=1;StatePersistenceHealthy=$true;MaintenanceHealthy=$true
        MaintenanceEnabled=$control.Enabled;MaintenanceRevision=$control.Revision }
}
try {
    $null=New-Item -ItemType Directory -Path (Join-Path $root 'Config') -Force
    Write-Fixture (Join-Path $root 'Config/Orchestrator-Cluster.json.txt') ([pscustomobject]@{ExpectedOrchestratorServers=@('SERVER-A');PeerHeartbeatStaleMinutes=5})
    $zero=Initialize-SmartM365OrchestratorMaintenance $root
    Write-Fixture (Join-Path $root 'SERVER-A/Orchestrator-Heartbeat.json.txt') (New-Heartbeat $zero)
    $incompatible=New-Heartbeat $zero; $incompatible.MaintenanceResetProtocol=0
    Write-Fixture (Join-Path $root 'SERVER-A/Orchestrator-Heartbeat.json.txt') $incompatible
    $refused=$false
    try { $null=Invoke-SmartM365CancelPendingWork -SharedDataFolderPath $root -Tenant test -ExpectedRevision 0 -Reason 'unsupported test' -AcknowledgementTimeoutSeconds 1 } catch { $refused=$true }
    if (-not $refused -or (Get-SmartM365OrchestratorMaintenanceState $root).Enabled) { throw 'An incompatible resident must be refused before pausing.' }
    Write-Fixture (Join-Path $root 'SERVER-A/Orchestrator-Heartbeat.json.txt') (New-Heartbeat $zero)
    $selection=[pscustomobject]@{SelectedJobs=@([pscustomobject]@{Name='AlreadyRunning'},[pscustomobject]@{Name='StillPending'})}
    $request=New-SmartM365OrchestratorPipelineRequest -SharedDataFolderPath $root -Tenant test -Pipeline Jobs -Selection $selection -ManifestHash 'fixture'
    $null=Set-SmartM365OrchestratorPipelineJobStatus -SharedDataFolderPath $root -BatchId $request.BatchId -JobName 'AlreadyRunning' -Status Running -OwnerServer 'SERVER-A'
    $orchestratorFolder=Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
    $guiAst=[Management.Automation.Language.Parser]::ParseFile((Join-Path $orchestratorFolder 'SmartM365-Inventory-Orchestrator-GUI.ps1'),[ref]$null,[ref]$null)
    $workerFunction=$guiAst.Find({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Get-GuiCancellationWork'},$true)
    . ([scriptblock]::Create($workerFunction.Extent.Text))
    $work=Get-GuiCancellationWork
    $progress=[hashtable]::Synchronized(@{Message='';LogError=''})
    $inputData=[pscustomobject]@{Kind='PendingWork';Phase='Validate';Root=$root;Tenant='test';Reason='offline test';Revision=[long]-1;LogPath='';ModulePaths=@(
        (Join-Path $orchestratorFolder 'SmartM365.Orchestrator.Management.psm1'),(Join-Path $orchestratorFolder 'SmartM365.Orchestrator.Pipeline.psm1'),
        (Join-Path $orchestratorFolder 'SmartM365.Orchestrator.Maintenance.psm1'),(Join-Path $orchestratorFolder 'SmartM365.Orchestrator.Insights.psm1'),
        (Join-Path $orchestratorFolder 'SmartM365.Orchestrator.PendingWork.psm1'))}
    $validation=& $work $progress $inputData
    if ($validation.Phase -ne 'PendingWorkValidate' -or $validation.Revision -ne 0 -or @($validation.BatchIds).Count -ne 1) { throw 'GUI cancellation preview did not enumerate the active request.' }
    $inputData.Tenant='wrong-tenant'; $refused=$false
    try { $null=& $work $progress $inputData } catch { $refused=$_.Exception.Message -match 'tenant' }
    if (-not $refused -or (Get-SmartM365OrchestratorMaintenanceState $root).Enabled) { throw 'GUI must preserve the preflight error and not pause the wrong tenant.' }
    $inputData.Tenant='test'
    $inputData.Phase='Publish'; $inputData.Revision=$validation.Revision
    $worker=Start-Job -ArgumentList $root -ScriptBlock {
        param($root)
        $controlPath=Join-Path $root 'Config/Orchestrator-Maintenance.json.txt'
        $heartbeatPath=Join-Path $root 'SERVER-A/Orchestrator-Heartbeat.json.txt'
        $end=[datetime]::UtcNow.AddSeconds(20)
        while ([datetime]::UtcNow -lt $end) {
            try {
                $control=Get-Content -LiteralPath $controlPath -Raw | ConvertFrom-Json
                if ($control.Enabled) {
                    $heartbeat=Get-Content -LiteralPath $heartbeatPath -Raw | ConvertFrom-Json
                    $heartbeat.MaintenanceEnabled=$true; $heartbeat.MaintenanceRevision=$control.Revision; $heartbeat.Timestamp=[datetime]::UtcNow.ToString('o')
                    $temp=$heartbeatPath+'.testtmp'; [IO.File]::WriteAllText($temp,($heartbeat | ConvertTo-Json -Depth 10),[Text.UTF8Encoding]::new($false))
                    [IO.File]::Move($temp,$heartbeatPath,$true)
                    break
                }
            } catch { }
            Start-Sleep -Milliseconds 100
        }
    }
    $published=& $work $progress $inputData
    if ($published.Phase -ne 'PendingWorkPublish') { throw 'GUI cancellation did not return the published result.' }
    $result=$published.Result
    if ($result.RunningJobsStopped -or $result.Revision -ne 2 -or (Get-SmartM365OrchestratorMaintenanceState $root).Enabled) { throw 'Reset did not resume correctly or claimed to stop running jobs.' }
    $requestAfter=Get-SmartM365OrchestratorPipelineRunStatus -SharedDataFolderPath $root -BatchId $request.BatchId
    if (-not $requestAfter.CancellationRequested -or @($requestAfter.Jobs | Where-Object { $_.JobName -eq 'StillPending' -and $_.Status -eq 'Cancelled' }).Count -ne 1 -or @($requestAfter.Jobs | Where-Object { $_.JobName -eq 'AlreadyRunning' -and $_.Status -eq 'Running' }).Count -ne 1) { throw 'Cancellation did not preserve running work and cancel pending request work.' }
    $hb=Get-Content (Join-Path $root 'SERVER-A/Orchestrator-Heartbeat.json.txt') -Raw | ConvertFrom-Json
    if (@($hb.RunningJobs).Count -ne 1) { throw 'The running job was not preserved in the synthetic heartbeat.' }
    $timeoutRoot=Join-Path $root 'TimeoutCase'
    $null=New-Item -ItemType Directory -Path (Join-Path $timeoutRoot 'Config') -Force
    Write-Fixture (Join-Path $timeoutRoot 'Config/Orchestrator-Cluster.json.txt') ([pscustomobject]@{ExpectedOrchestratorServers=@('SERVER-A');PeerHeartbeatStaleMinutes=5})
    $timeoutZero=Initialize-SmartM365OrchestratorMaintenance $timeoutRoot
    Write-Fixture (Join-Path $timeoutRoot 'SERVER-A/Orchestrator-Heartbeat.json.txt') (New-Heartbeat $timeoutZero)
    $failed=$false
    try { $null=Invoke-SmartM365CancelPendingWork -SharedDataFolderPath $timeoutRoot -Tenant test -ExpectedRevision 0 -Reason 'timeout test' -AcknowledgementTimeoutSeconds 1 } catch { $failed=$true }
    if (-not $failed -or -not (Get-SmartM365OrchestratorMaintenanceState $timeoutRoot).Enabled) { throw 'A missing acknowledgement must fail and leave the safe pause active.' }
    'PASS: coordinated pending-work cancellation leaves running jobs alone'
}
finally {
    if ($worker) { Stop-Job $worker -ErrorAction SilentlyContinue; Remove-Job $worker -Force -ErrorAction SilentlyContinue }
    if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force }
}
