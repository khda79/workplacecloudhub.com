#requires -Version 7.0
$ErrorActionPreference='Stop'
if ([Threading.Thread]::CurrentThread.ApartmentState -ne [Threading.ApartmentState]::STA) { throw 'Run with -STA.' }
$gui=Join-Path $PSScriptRoot '../SmartInventory/Orchestrator/SmartM365-Inventory-Orchestrator-GUI.ps1'
$result=. $gui -Demo -ValidateOnly -Tenant 'DO-NOT-LOAD' -SharedDataFolderPath 'Z:\DO-NOT-READ' -GuiLogFolderPath 'Z:\DO-NOT-WRITE'
$demoState=$script:OrchestratorDemo
$before=@($demoState.Data.Running).Count
$runningNames=@($demoState.Data.Running | ForEach-Object Job)
$recoverable=@($demoState.Data.Pending | Where-Object { $_.Job -notin $runningNames } | ForEach-Object Job | Sort-Object -Unique)
$requestsBefore=@($demoState.Data.Requests).Count
$demoState.Controls.RecoverJobsButton.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
if ($recoverable.Count -and (@($demoState.Data.Requests).Count -ne ($requestsBefore+1) -or @($demoState.Data.Requests[0].Rows | Where-Object { $_.JobName -in $runningNames }).Count)) { throw 'Demo recovery duplicated running work or omitted its request.' }
if (@($demoState.Data.Running).Count -ne $before) { throw 'Demo recovery changed running jobs.' }
$button=$demoState.Controls.CancelPendingWorkButton
$button.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
if (@($demoState.Data.Running).Count -ne $before -or @($demoState.Data.Pending).Count -ne 0) { throw 'Demo cancellation did not preserve running jobs and clear pending work.' }
if ($result.OutputRoot -and (Test-Path -LiteralPath $result.OutputRoot)) {
    $expected=[IO.Path]::GetFullPath((Join-Path ([IO.Path]::GetTempPath()) 'SmartM365-Demo/InventoryOrchestrator')).TrimEnd('\')+'\'
    $actual=[IO.Path]::GetFullPath([string]$result.OutputRoot)
    if (-not $actual.StartsWith($expected,[StringComparison]::OrdinalIgnoreCase)) { throw 'Unexpected demo output path.' }
    Remove-Item -LiteralPath $actual -Recurse -Force
}
'PASS: demo button preserves running jobs and cancels pending work'
