#requires -Version 7.0
$ErrorActionPreference='Stop'
$path=Join-Path $PSScriptRoot '../SmartInventory/Orchestrator/SmartM365-Inventory-Orchestrator.ps1'
$ast=[Management.Automation.Language.Parser]::ParseFile($path,[ref]$null,[ref]$null)
foreach($name in @('Update-OrchestratorRecoveryScheduleCursor','ConvertFrom-StateTime','ConvertTo-StateTime')) {
    $node=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)
    if(-not $node){throw "Missing $name"}; . ([scriptblock]::Create($node.Extent.Text))
}
$state=@{LastScheduledOccurrence='2026-11-12T00:15:00';LastStatus='Failed'}
Update-OrchestratorRecoveryScheduleCursor -State $state -LatestDue ([datetime]'2026-11-13T00:15:00')
if ((ConvertFrom-StateTime $state.LastScheduledOccurrence) -ne [datetime]'2026-11-13T00:15:00' -or $state.LastStatus -ne 'Failed') { throw 'Recovery cursor did not advance without inventing success.' }
Update-OrchestratorRecoveryScheduleCursor -State $state -LatestDue ([datetime]'2026-11-12T00:15:00')
if ((ConvertFrom-StateTime $state.LastScheduledOccurrence) -ne [datetime]'2026-11-13T00:15:00') { throw 'Recovery cursor moved backwards.' }
$start=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Start-InventoryJob'},$true).Extent.Text
if ($start.IndexOf('Update-OrchestratorRecoveryScheduleCursor') -lt $start.IndexOf('$process = Start-Process')) { throw 'Recovery cursor must not advance before process creation.' }
'PASS: recovery cursor advances monotonically only after launch'
