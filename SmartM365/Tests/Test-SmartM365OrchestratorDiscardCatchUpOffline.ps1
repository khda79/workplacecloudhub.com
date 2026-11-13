#requires -Version 7.0
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot '../SmartInventory/Orchestrator/SmartM365.Orchestrator.Maintenance.psm1') -Force
$root=Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-DiscardCatchUp-' + [guid]::NewGuid().ToString('N'))
function Write-Fixture($Path,$Value) {
    $null=New-Item -ItemType Directory -Path (Split-Path $Path -Parent) -Force
    [IO.File]::WriteAllText($Path,($Value | ConvertTo-Json -Depth 20),[Text.UTF8Encoding]::new($false))
}
function Write-Heartbeat($State,[bool]$Reset) {
    Write-Fixture (Join-Path $root 'SERVER-A/Orchestrator-Heartbeat.json.txt') ([pscustomobject]@{
        ScriptVersion='1.5.90';Timestamp=[datetime]::UtcNow.ToString('o');Lifecycle='Running';RunningJobs=@([pscustomobject]@{Name='Active'})
        MaintenanceProtocol=1;MaintenanceCatchUpProtocol=1;MaintenanceResetProtocol=[int]$Reset;MaintenanceHealthy=$true
        MaintenanceEnabled=$State.Enabled;MaintenanceRevision=$State.Revision
    })
}
try {
    $null=New-Item -ItemType Directory -Path (Join-Path $root 'Config') -Force
    Write-Fixture (Join-Path $root 'Config/Orchestrator-Cluster.json.txt') ([pscustomobject]@{ExpectedOrchestratorServers=@('SERVER-A');PeerHeartbeatStaleMinutes=5})
    Write-Fixture (Join-Path $root 'Config/Orchestrator-Jobs.json.txt') ([pscustomobject]@{Jobs=@([pscustomobject]@{Name='Example';Enabled=$true;AssignmentMode='Elected';Schedule=[pscustomobject]@{Type='Daily';Times=@('00:05');DaysOfWeek=@()}})})
    $zero=Initialize-SmartM365OrchestratorMaintenance $root
    Write-Heartbeat $zero $true
    $enabled=Set-SmartM365OrchestratorMaintenance $root $true 0 'first pause'
    Write-Heartbeat $enabled $true
    $resumed=Set-SmartM365OrchestratorMaintenance $root $false 1 'create catch-up' -CatchUpMissedOccurrences
    if (@($resumed.CatchUpWindows).Count -ne 1) { throw 'Fixture did not create a pending catch-up window.' }
    Write-Heartbeat $resumed $true
    $enabled=Set-SmartM365OrchestratorMaintenance $root $true 2 'second pause' -PendingWorkReset
    foreach ($origin in @('due','retry','pipeline','pipeline-retry','forced','maintenance-catch-up')) {
        if (Test-SmartM365OrchestratorMaintenanceLaunch $enabled $origin (Get-Date)) { throw "Pending-work reset must pause $origin launches." }
    }
    Write-Heartbeat $enabled $false
    $refused=$false
    try { $null=Set-SmartM365OrchestratorMaintenance $root $false 3 'reset' -DiscardPendingCatchUp } catch { $refused=$true }
    if (-not $refused -or -not (Get-SmartM365OrchestratorMaintenanceState $root).Enabled) { throw 'Unsupported resident must leave maintenance active.' }
    Write-Heartbeat $enabled $true
    $reset=Set-SmartM365OrchestratorMaintenance $root $false 3 'reset' -DiscardPendingCatchUp
    if ($reset.Enabled -or @($reset.CatchUpWindows).Count -ne 0) { throw 'Reset did not discard the pending window.' }
    $receipt=@(Get-ChildItem (Join-Path $root 'Audit/Orchestrator-MaintenanceReset') -File -Filter '*.json.txt')
    if ($receipt.Count -ne 1 -or @((Get-Content $receipt[0].FullName -Raw | ConvertFrom-Json).PreviousControl.CatchUpWindows).Count -ne 1) { throw 'Reset did not preserve the old control as an audit receipt.' }
    if (Test-SmartM365OrchestratorMaintenanceLaunch $reset 'maintenance-catch-up' (Get-Date)) { throw 'Discarded catch-up is still permitted.' }
    'PASS: audited catch-up reset and protocol gating'
}
finally { if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force } }
