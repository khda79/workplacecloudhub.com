#Requires -Version 7.0
<#
.SYNOPSIS
Offline WPF selection and asynchronous cancellation tests with synthetic delays.
.VERSION
1.1.0
#>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
$folder=Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
Import-Module (Join-Path $folder 'SmartM365.Orchestrator.GuiWorker.psm1') -Force
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $folder 'SmartM365-Inventory-Orchestrator-GUI.ps1'),[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'GUI parsing failed.'}
foreach($name in @('ConvertFrom-OrchestratorGuiXaml','Update-SelectedPipelineRequest','Set-GuiRequestRows','Refresh-RequestsView','Get-GuiCancellationWork','Set-GuiCancellationBusy','Start-GuiPipelineCancellation','Stop-SelectedPipelineRequest','Stop-AllPipelineRequests','Receive-GuiPipelineCancellation','Invoke-AutoRefresh')){
    $definition=$ast.Find({param($node)$node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name},$true)
    if($null -eq $definition){throw "Missing GUI function: $name"}
    . ([scriptblock]::Create($definition.Extent.Text))
}
$assignment=$ast.Find({param($node)$node -is [Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq '$xaml'},$true)
$window=ConvertFrom-OrchestratorGuiXaml -Text $assignment.Right.Expression.Value
$script:Controls=@{}
foreach($name in @('RequestsGrid','RequestJobsGrid','CancelRequestButton','CancelAllRequestsButton','CancellationReasonBox','CancellationProgressText','RefreshButton','ValidateButton','RebalanceButton','PublishButton','HistoryRefreshButton','RequestRunButton','EnableMaintenanceButton','DisableMaintenanceButton','RollbackButton')){
    $script:Controls[$name]=$window.FindName($name)
    if(-not $script:Controls[$name]){throw "Missing XAML control: $name"}
}
$root=Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-GuiCancellation-Test-'+[guid]::NewGuid().ToString('N'))
$null=New-Item -ItemType Directory -Path $root
$managementModulePath=Join-Path $root 'SyntheticCancellation.psm1'
# Mocks execute on the real worker thread, touching only this temporary folder.
$mock=@'
function Read-State {param($Root) Get-Content -LiteralPath (Join-Path $Root 'state.json.txt') -Raw|ConvertFrom-Json}
function Get-SmartM365OrchestratorActivePipelineRuns {
    param($SharedDataFolderPath)
    $state=Read-State $SharedDataFolderPath
    Start-Sleep -Milliseconds $state.DelayMs
    foreach($batch in $state.Active){[pscustomobject]@{BatchId=$batch}}
}
function Get-SmartM365OrchestratorPipelineRunStatus {
    param($SharedDataFolderPath,$BatchId)
    Start-Sleep -Milliseconds (Read-State $SharedDataFolderPath).DelayMs
    [pscustomobject]@{BatchId=$BatchId}
}
function Stop-SmartM365OrchestratorPipelineRequest {
    param($SharedDataFolderPath,$BatchId,$Tenant,$Reason,[switch]$ValidateOnly)
    $state=Read-State $SharedDataFolderPath
    Start-Sleep -Milliseconds $state.DelayMs
    $phase=if($ValidateOnly){'Preview'}else{'Publish'}
    [IO.File]::AppendAllText((Join-Path $SharedDataFolderPath 'calls.txt'),"$phase|$BatchId|$Tenant|$Reason`n")
    if($ValidateOnly -and $state.FailPreview){throw 'Synthetic readiness refusal'}
    if(-not $ValidateOnly -and $BatchId -eq $state.FailBatch){throw 'Synthetic late publication refusal'}
    [pscustomobject]@{BatchId=$BatchId;OverallStatus='Cancelling';CancelledCount=2;PendingCount=1}
}
function Get-SmartM365OrchestratorRecentPipelineRuns {
    param($SharedDataFolderPath,$Count)
    $state=Read-State $SharedDataFolderPath
    Start-Sleep -Milliseconds $state.DelayMs
    if($state.FailRefresh){throw 'Synthetic display refresh failure'}
    [pscustomobject]@{BatchId='BatchA';Status='Cancelling';JobRows=@([pscustomobject]@{JobName='Running';Status='Running'})}
}
function Write-SmartM365OrchestratorManagementLog {
    param($Path,$Message,$Level)
    [IO.File]::AppendAllText($Path,"$Level|$([Threading.Thread]::CurrentThread.ManagedThreadId)|$Message`n")
}
Export-ModuleMember -Function Get-SmartM365OrchestratorActivePipelineRuns,Get-SmartM365OrchestratorPipelineRunStatus,Stop-SmartM365OrchestratorPipelineRequest,Get-SmartM365OrchestratorRecentPipelineRuns,Write-SmartM365OrchestratorManagementLog
'@
[IO.File]::WriteAllText($managementModulePath,$mock,[Text.UTF8Encoding]::new($false))
$script:Cases=0;$script:Failures=@();$script:Activity=@();$script:RecentRows=@();$script:RequestRows=@()
$script:CancellationBusy=$false;$script:CancellationUiState=@{}
$script:CancellationContext=[pscustomobject]@{Worker=$null;Progress=[hashtable]::Synchronized(@{Message='';LogError=''});Input=$null;Clock=[Diagnostics.Stopwatch]::new()}
$script:SharedDataFolderPath=$root;$Tenant='test';$script:GuiLogPath=Join-Path $root 'activity.log'
$script:ConfirmResult=$true;$script:ConfirmationCalls=0;$script:AddFuture=$false;$script:ConfirmationMessage='';$script:UseRealProviders=$false
function Assert-Case {param([bool]$Condition,[string]$Message) $script:Cases++;if(-not $Condition){throw $Message}}
function Write-GuiException {param($Context,$ErrorRecord) $script:Failures+=$ErrorRecord.Exception.Message}
function Show-GuiCancellationError {param($Message) $script:Failures+=$Message}
function Write-GuiActivity {param($Level,$Message,[switch]$DisplayOnly) Assert-Case ([bool]$DisplayOnly) 'Persistent log on UI thread.';$script:Activity+=$Message}
function Get-SmartM365OrchestratorRecentPipelineRuns {param($SharedDataFolderPath,$Count) $script:RecentRows}
function Refresh-OperationsView {throw 'Automatic shared reads overlapped cancellation.'}
function Confirm-PipelineCancellation {
    param($Message,$Title)
    $script:ConfirmationCalls++;$script:ConfirmationMessage=$Message
    if($script:AddFuture){$script:State.Active+=@('FutureBatch');Save-State}
    $script:ConfirmResult
}
function Start-SmartM365OrchestratorGuiWorker {
    param($Context,$Work,$InputObject)
    if(-not $script:UseRealProviders){$InputObject.ModulePaths=@($managementModulePath)}
    SmartM365.Orchestrator.GuiWorker\Start-SmartM365OrchestratorGuiWorker -Context $Context -Work $Work -InputObject $InputObject
}
function Save-State {[IO.File]::WriteAllText((Join-Path $root 'state.json.txt'),($script:State|ConvertTo-Json),[Text.UTF8Encoding]::new($false))}
function Reset-Case {
    param([string[]]$Active=@('BatchA'),[int]$Delay=0,[bool]$FailPreview=$false,[string]$FailBatch='', [bool]$FailRefresh=$false)
    $script:State=@{Active=$Active;DelayMs=$Delay;FailPreview=$FailPreview;FailBatch=$FailBatch;FailRefresh=$FailRefresh};Save-State
    [IO.File]::WriteAllText((Join-Path $root 'calls.txt'),'')
    $script:Controls.CancellationReasonBox.Text='Synthetic test'
    $script:Failures=@();$script:Activity=@();$script:ConfirmationCalls=0;$script:ConfirmResult=$true;$script:AddFuture=$false
}
function Wait-GuiOperation {
    $frame=[Windows.Threading.DispatcherFrame]::new()
    $timer=[Windows.Threading.DispatcherTimer]::new();$timer.Interval=[timespan]::FromMilliseconds(20)
    $script:UiTicks=0;$script:WaitDeadline=[datetime]::UtcNow.AddSeconds(20)
    $tick={$script:UiTicks++;if(-not $script:CancellationBusy -or [datetime]::UtcNow -gt $script:WaitDeadline){$frame.Continue=$false}}
    $timer.Add_Tick($tick);$timer.Start()
    try{[Windows.Threading.Dispatcher]::PushFrame($frame)}finally{$timer.Stop();$timer.Remove_Tick($tick)}
    Assert-Case (-not $script:CancellationBusy) 'GUI operation exceeded offline bound.'
    Assert-Case ($null -eq $script:CancellationContext.Worker) 'Completed worker not disposed.'
}
$script:CancellationTimer=[Windows.Threading.DispatcherTimer]::new();$script:CancellationTimer.Interval=[timespan]::FromMilliseconds(40)
$script:CancellationTimer.Add_Tick({Receive-GuiPipelineCancellation});$script:CancellationTimer.Start()
$selection=$ast.Find({param($node)$node -is [Management.Automation.Language.InvokeMemberExpressionAst] -and $node.Expression.Extent.Text -eq '$script:Controls.RequestsGrid' -and $node.Member.Value -eq 'Add_SelectionChanged'},$true)
. ([scriptblock]::Create($selection.Extent.Text))
try{
    $rows=@(foreach($count in 0,1,2,4){[pscustomobject]@{BatchId="Batch$count";Status='Running';JobRows=@(foreach($i in 1..4|Select-Object -First $count){[pscustomobject]@{JobName="Job$i";Status='Pending'}})}})
    $script:Controls.RequestsGrid.ItemsSource=$rows
    foreach($row in $rows+$rows[1]+$rows[2]){
        $script:Controls.RequestsGrid.SelectedItem=$row
        Assert-Case ($script:Controls.RequestJobsGrid.ItemsSource -is [Collections.IEnumerable]) 'ItemsSource unwrapped.'
        Assert-Case ($script:Controls.RequestJobsGrid.Items.Count -eq $row.JobRows.Count) 'Job count changed.'
        Assert-Case $script:Controls.CancelRequestButton.IsEnabled 'Active selection disabled cancellation.'
    }
    $script:Controls.RequestJobsGrid.SelectedItems.Add($rows[2].JobRows[0])|Out-Null
    $script:Controls.RequestJobsGrid.SelectedItems.Add($rows[2].JobRows[1])|Out-Null
    Assert-Case ($script:Controls.RequestJobsGrid.SelectedItems.Count -eq 2) 'Two-job selection failed.'
    $script:Controls.RequestsGrid.SelectedItem=$null;Update-SelectedPipelineRequest
    Assert-Case ($script:Controls.RequestJobsGrid.Items.Count -eq 0 -and -not $script:Controls.CancelRequestButton.IsEnabled) 'Empty selection retained rows.'
    $script:RecentRows=@($rows[1]);Refresh-RequestsView
    $script:Controls.RequestsGrid.SelectedItem=$rows[1];Refresh-RequestsView
    Assert-Case ($script:Controls.RequestsGrid.SelectedItem.BatchId -eq 'Batch1' -and $script:Controls.RequestJobsGrid.Items.Count -eq 1) 'Singleton refresh lost selection.'
    $script:RecentRows=@();Refresh-RequestsView
    Assert-Case (-not $script:Controls.CancelAllRequestsButton.IsEnabled) 'Empty refresh left cancellation enabled.'
    Reset-Case;$script:Controls.CancellationReasonBox.Text=' '
    $failure='';try{Stop-AllPipelineRequests}catch{$failure=$_.Exception.Message}
    Assert-Case ($failure -eq 'Enter a cancellation reason.' -and -not $script:CancellationBusy) 'Empty reason started a worker.'
    Reset-Case -Active @();Stop-AllPipelineRequests;Wait-GuiOperation
    Assert-Case ([bool]($script:Failures -match 'No active pipeline requests remain.') -and $script:ConfirmationCalls -eq 0) 'Empty list opened confirmation.'
    Reset-Case -Active @('OutsideRecentHistory') -FailPreview $true;Stop-AllPipelineRequests;Wait-GuiOperation
    Assert-Case ([bool]($script:Failures -match 'Synthetic readiness refusal') -and $script:ConfirmationCalls -eq 0) 'Readiness refusal bypassed.'
    Assert-Case ([bool]((Get-Content (Join-Path $root 'calls.txt')) -match 'OutsideRecentHistory')) 'All-action used displayed history only.'
    Reset-Case -Active @('BatchA','BatchB');$script:ConfirmResult=$false;Stop-AllPipelineRequests;Wait-GuiOperation
    Assert-Case ($script:ConfirmationCalls -eq 1 -and -not ((Get-Content (Join-Path $root 'calls.txt')) -match '^Publish')) 'Declining published cancellation.'
    Assert-Case ($script:Controls.CancellationReasonBox.Text -eq 'Synthetic test') 'Declining lost reason.'
    Reset-Case -Active @('BatchA','BatchB') -Delay 350;$script:AddFuture=$true
    $script:Controls.RefreshButton.IsEnabled=$true;Stop-AllPipelineRequests
    $originalWorker=$script:CancellationContext.Worker;Stop-AllPipelineRequests
    Assert-Case ([object]::ReferenceEquals($originalWorker,$script:CancellationContext.Worker)) 'Double click replaced worker.'
    Assert-Case (-not $script:Controls.CancelAllRequestsButton.IsEnabled -and -not $script:Controls.RefreshButton.IsEnabled) 'Competing operations stayed enabled.'
    Update-SelectedPipelineRequest;Invoke-AutoRefresh;Refresh-RequestsView
    $clock=[Diagnostics.Stopwatch]::StartNew();$received=Receive-SmartM365OrchestratorGuiWorker -Context $script:CancellationContext;$clock.Stop()
    Assert-Case ($null -eq $received -and $clock.ElapsedMilliseconds -lt 150) 'Active-worker receive blocked UI.'
    Wait-GuiOperation
    Assert-Case ($script:UiTicks -gt 20) 'WPF dispatcher stopped during network delays.'
    Assert-Case ($script:ConfirmationCalls -eq 1 -and $script:ConfirmationMessage -match 'BatchA' -and $script:ConfirmationMessage -match 'BatchB' -and $script:ConfirmationMessage -match 'Running collectors continue') 'Confirmation lost scope or protection.'
    $calls=@(Get-Content (Join-Path $root 'calls.txt')|Where-Object {$_ -match '^Publish'})
    Assert-Case ($calls.Count -eq 2 -and -not ($calls -match 'FutureBatch')) 'Future batch or duplicate cancellation.'
    Assert-Case ($script:Failures.Count -eq 0 -and $script:Controls.CancellationReasonBox.Text -eq '' -and $script:Controls.RefreshButton.IsEnabled) 'Success did not restore controls.'
    $mainThread=[Threading.Thread]::CurrentThread.ManagedThreadId
    Assert-Case (-not ((Get-Content $script:GuiLogPath) -match "\|$mainThread\|")) 'Persistent log ran on WPF thread.'
    Reset-Case -Active @('BatchA','BatchB') -FailBatch 'BatchB';Stop-AllPipelineRequests;Wait-GuiOperation
    Assert-Case ([bool]($script:Failures -match 'Synthetic late publication refusal') -and $script:Activity.Count -eq 1 -and $script:Controls.CancellationReasonBox.Text -eq 'Synthetic test') 'Partial publication hidden or reason lost.'
    Reset-Case -FailRefresh $true;Stop-AllPipelineRequests;Wait-GuiOperation
    Assert-Case ($script:Failures.Count -eq 0 -and $script:Controls.CancellationProgressText.Text -match 'Display refresh failed') 'Refresh failure confused with cancellation failure.'
    Reset-Case;$script:Controls.RequestsGrid.SelectedItem=$script:RequestRows[0]
    Stop-SelectedPipelineRequest;Wait-GuiOperation
    Assert-Case ($script:ConfirmationCalls -eq 1 -and @(Get-Content (Join-Path $root 'calls.txt')|Where-Object {$_ -match '^Publish'}).Count -eq 1) 'Selected cancellation changed scope.'
    foreach($buttonName in @('CancelRequestButton','CancelAllRequestsButton')){
        $click=$ast.Find({param($node)$node -is [Management.Automation.Language.InvokeMemberExpressionAst] -and $node.Expression.Extent.Text -eq ('$script:Controls.'+$buttonName) -and $node.Member.Value -eq 'Add_Click'},$true)
        Assert-Case ($null -ne $click) "Button handler missing: $buttonName"
        . ([scriptblock]::Create($click.Extent.Text))
    }
    Reset-Case;$script:Controls.CancelAllRequestsButton.IsEnabled=$true
    $script:Controls.CancelAllRequestsButton.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Button]::ClickEvent));Wait-GuiOperation
    Assert-Case ($script:ConfirmationCalls -eq 1) 'Bulk click not wired.'
    $tab=$script:Controls.RequestsGrid.Parent.Parent.Parent;$tab.IsSelected=$true
    $layout=$window.Content;$layout.Measure([Windows.Size]::new(1180,720));$layout.Arrange([Windows.Rect]::new(0,0,1180,720));$layout.UpdateLayout()
    $button=$script:Controls.CancelAllRequestsButton;$position=$button.TranslatePoint([Windows.Point]::new(0,0),$layout)
    Assert-Case ($button.ActualWidth -gt 0 -and $position.X+$button.ActualWidth -lt 1180) 'Bulk button clipped.'
    Assert-Case ($script:Controls.CancellationProgressText.ActualHeight -gt 0) 'Progress not visible.'
    # Run the same asynchronous UI workload with real providers and synthetic state.
    # Pending/retry may change; started/completed jobs and audit scope must remain intact.
    $actualRoot=Join-Path $root 'ActualShared'
    $null=New-Item -ItemType Directory -Path (Join-Path $actualRoot 'Config') -Force
    $null=New-Item -ItemType Directory -Path (Join-Path $actualRoot 'WorkerA') -Force
    [IO.File]::WriteAllText((Join-Path $actualRoot 'Config/Orchestrator-Cluster.json.txt'),'{"ExpectedOrchestratorServers":["WorkerA"],"PeerHeartbeatStaleMinutes":5}')
    [IO.File]::WriteAllText((Join-Path $actualRoot 'WorkerA/Orchestrator-Heartbeat.json.txt'),(@{Timestamp=[datetime]::UtcNow.ToString('o');Lifecycle='Running';Tenant='test';PipelineCancellationProtocol=1;ScriptVersion='1.5.41'}|ConvertTo-Json))
    Import-Module (Join-Path $folder 'SmartM365.Orchestrator.Pipeline.psm1') -Force
    $selection=[pscustomobject]@{SelectedJobs=@(foreach($name in 'Waiting','Retry','Started','Finished'){[pscustomobject]@{Name=$name}})}
    $run=New-SmartM365OrchestratorPipelineRequest -SharedDataFolderPath $actualRoot -Tenant test -Pipeline Jobs -Selection $selection -ManifestHash 'synthetic'
    foreach($item in @(@('Retry','RetryScheduled'),@('Started','Running'),@('Finished','Success'))){
        $null=Set-SmartM365OrchestratorPipelineJobStatus -SharedDataFolderPath $actualRoot -BatchId $run.BatchId -JobName $item[0] -Status $item[1] -OwnerServer WorkerA
    }
    $requestHash=(Get-FileHash $run.RequestPath).Hash
    $protected=@(foreach($name in 'Started','Finished'){
        $file=Get-SmartM365OrchestratorPipelineJobStatusPath -SharedDataFolderPath $actualRoot -BatchId $run.BatchId -JobName $name
        [pscustomobject]@{Path=$file;Hash=(Get-FileHash $file).Hash}
    })
    $managementModulePath=(Resolve-Path (Join-Path $folder 'SmartM365.Orchestrator.Management.psm1')).Path
    $script:SharedDataFolderPath=$actualRoot;$script:UseRealProviders=$true
    $script:Controls.CancellationReasonBox.Text='Synthetic real-provider test';$script:ConfirmResult=$false
    $script:ConfirmationCalls=0;$script:Failures=@();Stop-AllPipelineRequests;Wait-GuiOperation
    Assert-Case ($script:Failures.Count -eq 0 -and $script:ConfirmationCalls -eq 1 -and (Get-FileHash $run.RequestPath).Hash -eq $requestHash) 'Real-provider preview or declined confirmation changed state.'
    $script:ConfirmResult=$true;Stop-AllPipelineRequests;Wait-GuiOperation
    $after=Get-SmartM365OrchestratorPipelineRunStatus -SharedDataFolderPath $actualRoot -BatchId $run.BatchId
    Assert-Case ($script:Failures.Count -eq 0 -and $after.OverallStatus -eq 'Cancelling' -and $after.CancelledCount -eq 2 -and $after.PendingCount -eq 1) 'Real async cancellation lost pending/running distinction.'
    foreach($item in $protected){Assert-Case ((Get-FileHash $item.Path).Hash -eq $item.Hash) 'Real async cancellation changed protected result bytes.'}
    Assert-Case ($after.Request.Cancellation.Reason -eq 'Synthetic real-provider test') 'Real cancellation audit lost reason.'
    Assert-Case ($script:RequestRows.Count -eq 1 -and $script:RequestRows[0].Status -eq 'Cancelling') 'Real background refresh failed to populate the grid.'
    $auditHash=(Get-FileHash $run.RequestPath).Hash
    $script:Controls.RequestsGrid.SelectedItem=$script:RequestRows[0]
    $script:Controls.CancellationReasonBox.Text='Synthetic idempotent click'
    $script:Controls.CancelRequestButton.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Button]::ClickEvent));Wait-GuiOperation
    Assert-Case ($script:Failures.Count -eq 0 -and (Get-FileHash $run.RequestPath).Hash -eq $auditHash) 'Selected button rewrote the audit of an already cancelled request.'
    $closing=$ast.Find({param($node)$node -is [Management.Automation.Language.InvokeMemberExpressionAst] -and $node.Member.Value -eq 'Add_Closing'},$true)
    Assert-Case ($null -ne $closing -and $closing.Extent.Text -match 'CancellationBusy' -and $closing.Extent.Text -match '\$eventArgs.Cancel = \$true') 'Active publication does not protect window closure.'
    "[{0}] PASS: {1} offline WPF/async request checks. No collectors, tenant calls or live shared writes." -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$script:Cases
}
finally{
    $script:CancellationTimer.Stop()
    if($script:CancellationContext.Worker){$null=$script:CancellationContext.Worker.Handle.AsyncWaitHandle.WaitOne(20000);$null=Receive-SmartM365OrchestratorGuiWorker -Context $script:CancellationContext}
    $window.Close()
    $resolved=[IO.Path]::GetFullPath($root)
    if($resolved.StartsWith([IO.Path]::GetFullPath([IO.Path]::GetTempPath()),[StringComparison]::OrdinalIgnoreCase) -and (Split-Path $resolved -Leaf) -like 'SmartM365-GuiCancellation-Test-*'){
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD1BaX/3NhbTCf7
# RYw/fnm4qpp3pT6/m0zAyLhnxe4viKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
# s0Q4yPEDH+JoMA0GCSqGSIb3DQEBCwUAME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91
# ZGh1Yi5jb20wHhcNMjYwNzEzMDgyMjM1WhcNMjkwNzEzMDgzMjI5WjBOMR4wHAYD
# VQQDDBV3b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRh
# Y3RAd29ya3BsYWNlY2xvdWRodWIuY29tMIIBojANBgkqhkiG9w0BAQEFAAOCAY8A
# MIIBigKCAYEAse6XztERSyHn9DVqj8Rdv0qjc5owqvgAIGaYxBmfiQuoM48Fo4Xt
# 1ovi9brLUtf55G4XgthNPCoanxfCRRg30IVRxaDfdPXJzYmgsM5tXlsuNU49lE7E
# PJk3+jEOgSCt8NKzmVPKpNRG0NmK0a8wm12cceYZOZlSYE0+ZtT6wy5PQQjMUqIx
# XnGjt4H0nfgZZa7D4FyARKOVg/Xr9sUq5jIn3zszvg4jjeb4b0DKJtfbHukhWc2Y
# oVFgswxVBXCWIaBnfF/cjqMfK/CaToT2trVb4hG4qcQ31s1nR4keoRaOw/vyd6ap
# rEtCsT22N/Jx0dz7fIo1tVyvIaVcHdN9LW3chn0en0OKZ6Ke1OH9wf2prl4KA6Ww
# VzrAZrOlXTAItdK7D9kKO/HeJd4PZvO53oy1LdmMGLSz3OLB9e5q7yo8rfqi5Ka9
# KzM2CrSzz1yphn/H90wz7Q2pm4FIlWdcj86A/0kmhYg+5Wqqbg1drrPXu4nEBwWN
# /dzoGtKZKHTdAgMBAAGjgZYwgZMwDgYDVR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoG
# CCsGAQUFBwMDMD8GA1UdEQQ4MDaBHWNvbnRhY3RAd29ya3BsYWNlY2xvdWRodWIu
# Y29tghV3b3JrcGxhY2VjbG91ZGh1Yi5jb20wDAYDVR0TAQH/BAIwADAdBgNVHQ4E
# FgQUXIOOADQM78XfPAncirgCECedg9gwDQYJKoZIhvcNAQELBQADggGBADhZUB2R
# 5J/Jw030xodhEWeCQ0vnJRaiEsjOxuArQREKH3lCrQ3UsUVl292d6LnQUSTH/jF7
# rovEZ+JN2GQ/LCrXRaCuwCEGZKzlSEbtYWhfwDyj6GpIPq8Y4SeXyjdq4/rrI1bm
# iTK4Sq7EoBlGJuX6l2nfvx1tTioSr11FoDfllJR7EYawRj9hBFJ0gG0b2SuYZMgW
# gaDKefcnJDmOwcRNAZUII0ss8EeyANukWSkNN5ILZ+iKDpQgZxgDLPTiRguCyx45
# PI5wrVTjV/pR7IrtSIfq8UladlrSZJyyDn3NV2ATvIZ6wNxbTmPFcE0uMg/EYzwd
# Tek+CgXL3TxUKeldJM4YDWPimNBRhOPXzBDiOQIj6WNswt/KM1oDLnA00CNtciPN
# dn+dXlneMvTEUah9wyt8o8tkLpoBw+KN+Bq/K0O1qPtS7umi70l45pPiej+mwbwq
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBY0w
# ggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFowDQYJKoZIhvcNAQEMBQAwZTELMAkG
# A1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRp
# Z2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENB
# MB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIzNTk1OVowYjELMAkGA1UEBhMCVVMx
# FTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNv
# bTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+RdSjwwIjBpM+zCpyUuySE98orY
# WcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20dq7J58soR0uRf1gU8Ug9SH8ae
# FaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7fgvMHhOZ0O21x4i0MG+4g1ckg
# HWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRAX7F6Zu53yEioZldXn1RYjgwr
# t0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raRmECQecN4x7axxLVqGDgDEI3Y
# 1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzUvK4bA3VdeGbZOjFEmjNAvwjX
# WkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2mHY9WV1CdoeJl2l6SPDgohIb
# Zpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkrfsCUtNJhbesz2cXfSwQAzH0c
# lcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaAsPvoZKYz0YkH4b235kOkGLim
# dwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxfjT/JvNNBERJb5RBQ6zHFynIW
# IgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEexcCPorF+CiaZ9eRpL5gdLfXZ
# qbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFOzX
# 44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6enIZ3z
# bcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGG
# GGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2Nh
# Y2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDBF
# BgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1UdIAQKMAgwBgYEVR0gADANBgkqhkiG
# 9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22Ftf3v1cHvZqsoYcs7IVeqRq7IviH
# GmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih9/Jy3iS8UgPITtAq3votVs/59Pes
# MHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYDE3cnRNTnf+hZqPC/Lwum6fI0POz3
# A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c2PR3WlxUjG/voVA9/HYJaISfb8rb
# II01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88nq2x2zm8jLfR+cWojayL/ErhULSd+
# 2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBrQwggScoAMCAQICEA3HrFcF
# /yGZLkBDIgw6SYYwDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTATBgNV
# BAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8G
# A1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTI1MDUwNzAwMDAwMFoX
# DTM4MDExNDIzNTk1OVowaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0
# LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFtcGlu
# ZyBSU0E0MDk2IFNIQTI1NiAyMDI1IENBMTCCAiIwDQYJKoZIhvcNAQEBBQADggIP
# ADCCAgoCggIBALR4MdMKmEFyvjxGwBysddujRmh0tFEXnU2tjQ2UtZmWgyxU7UNq
# EY81FzJsQqr5G7A6c+Gh/qm8Xi4aPCOo2N8S9SLrC6Kbltqn7SWCWgzbNfiR+2fk
# HUiljNOqnIVD/gG3SYDEAd4dg2dDGpeZGKe+42DFUF0mR/vtLa4+gKPsYfwEu7EE
# bkC9+0F2w4QJLVSTEG8yAR2CQWIM1iI5PHg62IVwxKSpO0XaF9DPfNBKS7Zazch8
# NF5vp7eaZ2CVNxpqumzTCNSOxm+SAWSuIr21Qomb+zzQWKhxKTVVgtmUPAW35xUU
# FREmDrMxSNlr/NsJyUXzdtFUUt4aS4CEeIY8y9IaaGBpPNXKFifinT7zL2gdFpBP
# 9qh8SdLnEut/GcalNeJQ55IuwnKCgs+nrpuQNfVmUB5KlCX3ZA4x5HHKS+rqBvKW
# xdCyQEEGcbLe1b8Aw4wJkhU1JrPsFfxW1gaou30yZ46t4Y9F20HHfIY4/6vHespY
# MQmUiote8ladjS/nJ0+k6MvqzfpzPDOy5y6gqztiT96Fv/9bH7mQyogxG9QEPHrP
# V6/7umw052AkyiLA6tQbZl1KhBtTasySkuJDpsZGKdlsjg4u70EwgWbVRSX1Wd4+
# zoFpp4Ra+MlKM2baoD6x0VR4RjSpWM8o5a6D8bpfm4CLKczsG7ZrIGNTAgMBAAGj
# ggFdMIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBTvb1NK6eQGfHrK
# 4pBW9i/USezLTjAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNV
# HQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEEazBp
# MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUH
# MAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRS
# b290RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0
# LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYGZ4EM
# AQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAF877FoAc/gc9EXZx
# ML2+C8i1NKZ/zdCHxYgaMH9Pw5tcBnPw6O6FTGNpoV2V4wzSUGvI9NAzaoQk97fr
# PBtIj+ZLzdp+yXdhOP4hCFATuNT+ReOPK0mCefSG+tXqGpYZ3essBS3q8nL2UwM+
# NMvEuBd/2vmdYxDCvwzJv2sRUoKEfJ+nN57mQfQXwcAEGCvRR2qKtntujB71WPYA
# gwPyWLKu6RnaID/B0ba2H3LUiwDRAXx1Neq9ydOal95CHfmTnM4I+ZI2rVQfjXQA
# 1WSjjf4J2a7jLzWGNqNX+DF0SQzHU0pTi4dBwp9nEC8EAqoxW6q17r0z0noDjs6+
# BFo+z7bKSBwZXTRNivYuve3L2oiKNqetRHdqfMTCW/NmKLJ9M+MtucVGyOxiDf06
# VXxyKkOirv6o02OoXN4bFzK0vlNMsvhlqgF2puE6FndlENSmE+9JGYxOGLS/D284
# NHNboDGcmWXfwXRy4kbu4QFhOm0xJuF2EZAOk5eCkhSxZON3rGlHqhpB/8MluDez
# ooIs8CVnrpHMiD2wL40mm53+/j7tFaxYKIqL0Q4ssd8xHZnIn/7GELH3IdvG2XlM
# 9q7WP/UwgOkw/HQtyRN62JK4S1C8uw3PdBunvAZapsiI5YKdvlarEvf8EA+8hcpS
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAIT9wzT35FTtvDD4/5
# khg1MA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjYwODA1MDAwMDAwWhcN
# MzcxMTA0MjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNiAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEAtnum8sn+zUr41JtMZbP9OMYw+HwJDpG5xkIu/lqcfNYmMX81YmsUiHLbh9yk
# peWBGKTLhYBrAN9Tdg/QEzG32XcObmgIblnr0CoQ3WSAeDZ6nH6X6VkFyYkJw3QB
# JREwvm4UhLzSxmwPA7cFKRTEOMsmEEj6qJk/dqLEAL+oQYuOwE2UuiX1Vnul8YRe
# IyWd4kgLn9gq6LNXM0UplkR6jL/QHxmb6fMoGBJYbnaUI7XD6cKDpekK2SVMld4i
# DbzeHDtOaaxldH5IxuNusQ69nd8/ZXEiB5Hbxj3RlK13cX1W4DlFXKdv/CEhM8Cj
# 1vvlmvhNroyPdRGbbpBlgyf8Wdu5N6ByhFwURn0U6ozlPoxN22v+fviUhP+6DR54
# 7OZnpBMWDfei1f5sVGwiiW/KQTWOK97g+4RJpPzPNV4VYMAwO2jM2Aty2QYPVmOQ
# TJm0msuXnJrSbl2gf9JylpkJlWXqk1Q4LJsxz+TELoQCZIljbgvTJgoPU2R12ydv
# 8i1UqL/adelA0y7U9Pmmtbze9Xx3rtajC5SzQd1jgfwAwsa90v9YcSPdmeoyoBBA
# /27cCL237l5DTYYPDLQ4ON3OLTGWnvRb6jDrf/T75gMRfUzSLCBQfBusm9+mSWRl
# C/Df6S/e9Q8i13CuhzOT2Jx+V/nlbXM4QoBwlUAhelwwJT0CAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFBTJY4owLtRK+26U8+bjQH717M3iMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAI3FOmEenVIK35ms
# CYB+fShAsWvSYvLBItoNdAgQ2jIqrGsVsluXMJU/+mRebBc52s6lbKAvOVPXaizm
# KkMLLflEEKDZQx4CkS2t8aHPjkXha3hYZ010htFa3dhNgmalH5vuWvh3tTCf4frT
# S7gPtGc4Z/xaPhQ2AB1mR8eEe/WbH0RWHvVIl6VwQ3+g5FKNfN2N/DWJkf13w2H+
# 2GfqEfbd35Ww8CvoYBjLNIDTadcPWdgsjsiOaK/7EsKJgLjUNIVgvcaFOLLQ/Glr
# A+0ZHJoFUbOr5SJN8zykPspXIXlpDJY/gqFUZRROeab9GVgmhbdOJcD/63RhxPah
# FUGbckRONqMe6DYAv6/mOG0pWd3cPStsdcS7buj5DyniwRY8yooMH6ptx5vpP/pZ
# zBPBeZD2U4IsthyxB5Jaa8qrOkB5z160TXiM5ADMspZ0TfD9MJoq0tFpFPssKRFh
# WeEDYPvcUuN7U7lvcdHl4ezQ3NT/7Ffs1sR1yh/LRbdZ3B3Vc6q2WmD8mDC0p9kz
# l2o73iVtS946IkEj7FkRsZGww1teYxERROC745xrtjvcw9ZyyUjHZWGRIpJeMNsP
# quCDf0fkyHtB+J4AiNZqCQk23rxh+KbpyMTNVKItJ5l92Svl20U9NbqMBOVYl1h5
# 4NEYLJq1/xHWFKPNK903zJZA9P2DMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIONME2kBURuMKHE6z/EuQeqDWOjzavr3fPWlhU7/tEH7MA0GCSqG
# SIb3DQEBAQUABIIBgI2V/nvTQZf8SStu0pz9oqiKE7Cm2cxeYd0CXD0U09uyx4hF
# JYC0JD+oGWOnUl+iEYpPFY+UoejGhuMXFvsrvQR5LP1sWXgsep2fdkKjmBFWlY0H
# FtElivkV9Y7sI3KU5wTsk+epHH6HOd3ynskPhbBTnQfj7A3lisdhvf8DO4SUMQfx
# kHPHJGhPbboav808A6XCo9P6PS+YwJKSbmyhhtUMvoQUZiSnlbXPgqy7yH9bA3Pg
# syz7lWEUpGUaqdgG1M1Zrjb02aG1RarcaBNeApE1HNSsW5z2tt9HnoL1LEjR861b
# D7ATU5et6S69SBtDL01yS/1UI6YXluTXVT4d6mk4R9J01SSXBuJHw0N4zCmnJhdb
# RdUr1qM565200w9GpoJXm+BImJ7eYw6vCq349C9nlmvGlvO0cS5ndFEFNiuXlfps
# GdbkMnsF6CsdpM9Xxvj/PpryNqDteVfzrN3POVR/Pg/lCq7i/BktlogROLmRJbJR
# lGfiDYsCq+LSrQxxdqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDUyMDQ4
# MzlaMC8GCSqGSIb3DQEJBDEiBCAaX6lzkJSu0j/KzB6YnC1bUXqeVVTD2EXJ38w3
# p37a0TANBgkqhkiG9w0BAQEFAASCAgA4kh3nx2//9wTRZm5hXOTz5ieBX6Z1tXu/
# wMK3rBNLRS8nUM2sDJpNOZ1/V463vwkloidQvpDYhXKPqZyGs8I9jv1Pu5Z+ct52
# wowFGDJZz08tbMXnM4yXfposOM6+0+3uPOVdXBcN+y/xOpuDRXmIlGXCUQXclb2M
# W9Hg0k9gIU7EWJnvYBBToefUVznkdZ1ilxPJEVRh9ik5TL2M4dYkIG7bIsAcr8xg
# SftMMwbGLewozaFQX7glLWDpKG+/kliRqeLw2AeMCo1yyqZKvUyNTUQGRlgJIqhJ
# cy3nxvL2o4xP3PiPhauBaeTRzyzqqlg3OgTqZfZTRclHP3kvcNaJIzoi325RRpkV
# gknSw+dKgZOLJIfaQDuDDmb9JgQdmmtZumQBuihnBz0/D3DKvK0wqhB/EG3yMTYQ
# sPDr8tqeZR8wiFqV0wItNCmhHBbAianPpcIYgiIvzPJKixNqIPErDWq6oKUUhJj9
# Bu1ynAKWZAvXzgkwexFigLUqpxXgkuuC4LP4rybvW6dXjFCDMjgS4yJe28JVtwHr
# d8XrlvnexHePvC33ctgKYHkprFd9+9QOQz/kLOkv1ASK4wiiX1f9C8XtMJ8uXwdz
# qGrVKlja6wesSDwuJEOoVcDk6/vKu1gYAqb19adMNNbxa1x1rroSJIEaBTu7Y2Ug
# EcpxtgObKw==
# SIG # End signature block
