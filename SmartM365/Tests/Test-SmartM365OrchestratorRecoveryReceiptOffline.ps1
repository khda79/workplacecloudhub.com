$ErrorActionPreference='Stop'
$base=Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
Import-Module (Join-Path $base 'SmartM365.Orchestrator.Pipeline.psm1') -Force
$root=Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-RecoveryReceipt-' + [guid]::NewGuid().ToString('N'))
try {
    $null=New-Item -ItemType Directory -Path $root -Force
    $due=[datetime]::UtcNow.AddHours(-2)
    if ($null -ne (Get-SmartM365OrchestratorRecoveryScheduleReceipt -SharedDataFolderPath $root -Tenant 'test' -JobName Example)) { throw 'Unexpected receipt.' }
    Set-SmartM365OrchestratorRecoveryScheduleReceipt -SharedDataFolderPath $root -Tenant 'test' -JobName Example -BatchId 'recovery-A' -CoveredThrough $due
    # A new owner reads the same durable admission even when the original attempt failed.
    $fromPeer=Get-SmartM365OrchestratorRecoveryScheduleReceipt -SharedDataFolderPath $root -Tenant 'test' -JobName Example
    if ($fromPeer.CoveredThroughUtc -ne $due -or $fromPeer.BatchId -ne 'recovery-A') { throw 'Peer did not read the shared cutoff.' }
    Set-SmartM365OrchestratorRecoveryScheduleReceipt -SharedDataFolderPath $root -Tenant 'test' -JobName Example -BatchId 'recovery-A' -CoveredThrough $due.AddDays(1)
    if ((Get-SmartM365OrchestratorRecoveryScheduleReceipt -SharedDataFolderPath $root -Tenant 'test' -JobName Example).CoveredThroughUtc -ne $due) { throw 'A retry consumed a future schedule.' }
    Set-SmartM365OrchestratorRecoveryScheduleReceipt -SharedDataFolderPath $root -Tenant 'test' -JobName Example -BatchId 'older' -CoveredThrough $due.AddHours(-1)
    if ((Get-SmartM365OrchestratorRecoveryScheduleReceipt -SharedDataFolderPath $root -Tenant 'test' -JobName Example).BatchId -ne 'recovery-A') { throw 'Receipt moved backwards.' }
    $later=$due.AddDays(1)
    Set-SmartM365OrchestratorRecoveryScheduleReceipt -SharedDataFolderPath $root -Tenant 'test' -JobName Example -BatchId 'recovery-B' -CoveredThrough $later
    if ((Get-SmartM365OrchestratorRecoveryScheduleReceipt -SharedDataFolderPath $root -Tenant 'test' -JobName Example).CoveredThroughUtc -ne $later) { throw 'Later receipt did not advance.' }
    $refused=$false; try { $null=Get-SmartM365OrchestratorRecoveryScheduleReceipt -SharedDataFolderPath $root -Tenant 'wrong' -JobName Example } catch { $refused=$true }; if (-not $refused) { throw 'Tenant mismatch accepted.' }
    $receiptPath=Join-Path $root 'RecoveryScheduleReceipts/Example.json.txt'
    [IO.File]::WriteAllText($receiptPath,'{ invalid json')
    $refused=$false; try { $null=Get-SmartM365OrchestratorRecoveryScheduleReceipt -SharedDataFolderPath $root -Tenant 'test' -JobName Example } catch { $refused=$true }; if (-not $refused) { throw 'Corrupt receipt accepted.' }
    $ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $base 'SmartM365-Inventory-Orchestrator.ps1'),[ref]$null,[ref]$null)
    $start=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Start-InventoryJob'},$true).Extent.Text
    if ($start.IndexOf('Set-SmartM365OrchestratorRecoveryScheduleReceipt') -lt 0 -or $start.IndexOf('Set-SmartM365OrchestratorRecoveryScheduleReceipt') -gt $start.IndexOf('$process = Start-Process')) { throw 'Receipt must precede process creation.' }
    $launch=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Invoke-LaunchPhase'},$true).Extent.Text
    if ($launch -notmatch "Get-SmartM365OrchestratorRecoveryScheduleReceipt" -or $launch -notmatch 'occurrence.ToUniversalTime\(\) -le \$receipt.CoveredThroughUtc') { throw 'Launch phase does not enforce shared cutoff.' }
    'PASS: shared recovery receipt survives failover and never asserts success'
} finally { if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force } }
