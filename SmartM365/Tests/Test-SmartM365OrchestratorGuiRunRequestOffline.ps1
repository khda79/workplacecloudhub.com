#Requires -Version 7.0
<#
.SYNOPSIS
Offline WPF tests of the asynchronous run request submission: wait window, slow, success,
error, uncertain result, double click and auto-refresh during the wait. No job is launched.
.VERSION
1.0.2
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
$guiPath=Join-Path $folder 'SmartM365-Inventory-Orchestrator-GUI.ps1'
$ast=[Management.Automation.Language.Parser]::ParseFile($guiPath,[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'GUI parsing failed.'}
foreach($name in @('ConvertFrom-OrchestratorGuiXaml','Update-SelectedPipelineRequest','Set-GuiRequestRows','Refresh-RequestsView','Invoke-AutoRefresh','Start-GuiPipelineCancellation',
    'Request-SelectedJobRun','Get-GuiJobRunSubmissionWork','Show-GuiRequestSubmitDialog','Close-GuiRequestSubmitDialog','Set-GuiRequestSubmitBusy','Start-GuiJobRunSubmission','Receive-GuiJobRunSubmission')){
    $definition=$ast.Find({param($node)$node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name},$true)
    if($null -eq $definition){throw "Missing GUI function: $name"}
    . ([scriptblock]::Create($definition.Extent.Text))
}
$assignment=$ast.Find({param($node)$node -is [Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq '$xaml'},$true)
$window=ConvertFrom-OrchestratorGuiXaml -Text $assignment.Right.Expression.Value
$script:Controls=@{}
foreach($name in @('RequestRunButton','RequestJobButton','RefreshButton','CancelRequestButton','CancelAllRequestsButton','RequestsGrid','RequestJobsGrid','StatusText','PlanningGrid',
    'IncludeDependenciesCheck','LaunchJobsGrid','LaunchIncludeDependenciesCheck','CancellationReasonBox')){
    $script:Controls[$name]=$window.FindName($name)
    if(-not $script:Controls[$name]){throw "Missing XAML control: $name"}
}
$root=Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-GuiRunRequest-Test-'+[guid]::NewGuid().ToString('N'))
$null=New-Item -ItemType Directory -Path $root
$mockModulePath=Join-Path $root 'SyntheticRunRequest.psm1'
# Mocks run on the real worker thread and only touch this temporary folder.
$mock=@'
function Read-State {param($Root) Get-Content -LiteralPath (Join-Path $Root 'state.json.txt') -Raw|ConvertFrom-Json}
function Get-SmartM365OrchestratorConfigurationPaths {param($SharedDataFolderPath) [pscustomobject]@{JobsPath=(Join-Path $SharedDataFolderPath 'jobs.json')}}
function Get-SmartM365JsonReadPath {param($Path,[switch]$Optional) if((Read-State (Split-Path $Path -Parent)).MissingJobs){$null}else{$Path}}
function New-SmartM365OrchestratorJobRunRequest {
    param($SharedDataFolderPath,$JobsPath,$JobName,$Tenant,[switch]$IncludeDependencies)
    $state=Read-State $SharedDataFolderPath
    [IO.File]::AppendAllText((Join-Path $SharedDataFolderPath 'calls.txt'),"$($JobName -join ',')|$Tenant|$([bool]$IncludeDependencies)|$([Threading.Thread]::CurrentThread.ManagedThreadId)`n")
    Start-Sleep -Milliseconds $state.DelayMs
    if($state.Fail -eq 'Before'){throw 'Synthetic refusal before writing'}
    $batch='Batch'+(Get-Date -Format 'HHmmssfff')
    $null=New-Item -ItemType Directory -Path (Join-Path (Join-Path $SharedDataFolderPath 'PipelineRuns') $batch) -Force
    if($state.Fail -eq 'After'){throw 'Synthetic failure after writing'}
    [pscustomobject]@{BatchId=$batch;TotalCount=2;QueuePosition=[int]$state.Queue;QueuedBehind=@(if($state.Queue){'Active1'})}
}
function Get-SmartM365OrchestratorRecentPipelineRuns {
    param($SharedDataFolderPath,$Count)
    if((Read-State $SharedDataFolderPath).FailRefresh){throw 'Synthetic display refresh failure'}
    foreach($folder in @(Get-ChildItem -LiteralPath (Join-Path $SharedDataFolderPath 'PipelineRuns') -Directory)){[pscustomobject]@{BatchId=$folder.Name;Status='Running';Created='';Pending=2;JobRows=@()}}
}
function Write-SmartM365OrchestratorManagementLog {param($Path,$Message,$Level) [IO.File]::AppendAllText($Path,"$Level|$([Threading.Thread]::CurrentThread.ManagedThreadId)|$Message`n")}
Export-ModuleMember -Function *
'@
[IO.File]::WriteAllText($mockModulePath,$mock,[Text.UTF8Encoding]::new($false))
$managementModulePath=$mockModulePath
$script:Cases=0;$script:RequestRows=@();$script:Activity=@();$script:Results=@();$script:Exceptions=@()
$script:CancellationBusy=$false;$script:RequestSubmitBusy=$false;$script:RequestSubmitUiState=@{};$script:RequestSubmitDialog=$null;$script:RequestSubmitDialogClosing=$false
$script:RequestSubmitContext=[pscustomobject]@{Worker=$null;Progress=[hashtable]::Synchronized(@{Message='';Stage='';LogError=''});Input=$null;Clock=[Diagnostics.Stopwatch]::new()}
$script:SharedDataFolderPath=$root;$Tenant='test';$script:GuiLogPath=Join-Path $root 'activity.log'
$script:ConfirmResult=$true;$script:ConfirmCalls=0;$script:RefreshCalls=0;$script:SharedReads=0;$script:ModuleOverride=$null
function Assert-Case {param([bool]$Condition,[string]$Message) $script:Cases++;if(-not $Condition){throw "FAIL: $Message"}}
function Write-GuiException {param($Context,$ErrorRecord) $script:Exceptions+="$Context : $($ErrorRecord.Exception.Message)"}
function Write-GuiActivity {param($Message,$Level='INFO',[switch]$DisplayOnly) Assert-Case ([bool]$DisplayOnly) 'Persistent log written on the WPF thread.';$script:Activity+="$Level|$Message"}
function Show-GuiJobRunRequestResult {param($Message,$Title,$Icon) $script:Results+=[pscustomobject]@{Message=$Message;Title=$Title;Icon=$Icon;DialogOpen=($null -ne $script:RequestSubmitDialog);Busy=$script:RequestSubmitBusy}}
function Confirm-GuiJobRunRequest {param($Message) $script:ConfirmCalls++;$script:LastConfirmMessage=$Message;$script:ConfirmResult}
function Test-JobDraftChanged {param($JobName) $false}
function Start-GuiRefresh {param($Mode) $script:RefreshCalls++}
function Get-SmartM365OrchestratorRecentPipelineRuns {param($SharedDataFolderPath,$Count) $script:SharedReads++;@()}
function Start-SmartM365OrchestratorGuiWorker {
    param($Context,$Work,$InputObject)
    $InputObject.ModulePaths=if($script:ModuleOverride){@($script:ModuleOverride)}else{@($mockModulePath)}
    SmartM365.Orchestrator.GuiWorker\Start-SmartM365OrchestratorGuiWorker -Context $Context -Work $Work -InputObject $InputObject
}
function Save-State {param([hashtable]$State) [IO.File]::WriteAllText((Join-Path $root 'state.json.txt'),($State|ConvertTo-Json),[Text.UTF8Encoding]::new($false))}
function Reset-Case {
    param([int]$Delay=0,[string]$Fail='',[bool]$FailRefresh=$false,[bool]$MissingJobs=$false,[int]$Queue=0)
    Save-State @{DelayMs=$Delay;Fail=$Fail;FailRefresh=$FailRefresh;MissingJobs=$MissingJobs;Queue=$Queue}
    [IO.File]::WriteAllText((Join-Path $root 'calls.txt'),'')
    $runs=Join-Path $root 'PipelineRuns';if(Test-Path $runs){Remove-Item -LiteralPath $runs -Recurse -Force}
    $script:Activity=@();$script:Results=@();$script:Exceptions=@();$script:ConfirmCalls=0;$script:ConfirmResult=$true;$script:RefreshCalls=0;$script:SharedReads=0;$script:ModuleOverride=$null
    $script:RequestRows=@();$script:Controls.RequestsGrid.ItemsSource=@()
    foreach($name in 'RequestRunButton','RequestJobButton','RefreshButton'){$script:Controls[$name].IsEnabled=$true}
}
function Get-Calls {@(Get-Content -LiteralPath (Join-Path $root 'calls.txt')|Where-Object {$_})}
function Wait-Submission {
    param([scriptblock]$During={},[scriptblock]$Later={})
    $frame=[Windows.Threading.DispatcherFrame]::new()
    $timer=[Windows.Threading.DispatcherTimer]::new();$timer.Interval=[timespan]::FromMilliseconds(20)
    $script:UiTicks=0;$script:WaitDeadline=[datetime]::UtcNow.AddSeconds(30);$script:DuringDone=$false;$script:LaterDone=$false
    $tick={
        $script:UiTicks++
        if(-not $script:DuringDone -and $script:UiTicks -ge 5){$script:DuringDone=$true;& $During}
        if(-not $script:LaterDone -and $script:UiTicks -ge 25 -and $script:RequestSubmitBusy){$script:LaterDone=$true;& $Later}
        if(-not $script:RequestSubmitBusy -or [datetime]::UtcNow -gt $script:WaitDeadline){$frame.Continue=$false}
    }
    $timer.Add_Tick($tick);$timer.Start()
    try{[Windows.Threading.Dispatcher]::PushFrame($frame)}finally{$timer.Stop();$timer.Remove_Tick($tick)}
    Assert-Case (-not $script:RequestSubmitBusy) 'Submission exceeded the offline bound.'
    Assert-Case ($null -eq $script:RequestSubmitContext.Worker) 'Completed worker not disposed.'
}
function Get-BarSnapshot {
    param($Track)
    $width=[Math]::Max(1,[int]$Track.ActualWidth);$height=[Math]::Max(1,[int]$Track.ActualHeight)
    # VisualBrush: RenderTargetBitmap of a nested element would apply its parent offset.
    $visual=[Windows.Media.DrawingVisual]::new();$context=$visual.RenderOpen()
    $context.DrawRectangle([Windows.Media.VisualBrush]::new($Track),$null,[Windows.Rect]::new(0,0,$width,$height));$context.Close()
    $bitmap=[Windows.Media.Imaging.RenderTargetBitmap]::new($width,$height,96,96,[Windows.Media.PixelFormats]::Pbgra32)
    $bitmap.Render($visual)
    $pixels=[byte[]]::new($width*$height*4);$bitmap.CopyPixels($pixels,$width*4,0)
    [Convert]::ToBase64String([Security.Cryptography.SHA256]::HashData($pixels))
}
$script:SubmitTimer=[Windows.Threading.DispatcherTimer]::new();$script:SubmitTimer.Interval=[timespan]::FromMilliseconds(40)
$script:SubmitTimer.Add_Tick({Receive-GuiJobRunSubmission});$script:SubmitTimer.Start()
foreach($buttonName in 'RequestRunButton','RequestJobButton'){
    $click=$ast.Find({param($node)$node -is [Management.Automation.Language.InvokeMemberExpressionAst] -and $node.Expression.Extent.Text -eq ('$script:Controls.'+$buttonName) -and $node.Member.Value -eq 'Add_Click'},$true)
    Assert-Case ($null -ne $click -and $click.Extent.Text -match 'Request-SelectedJobRun') "Request entry not wired through Request-SelectedJobRun: $buttonName"
    . ([scriptblock]::Create($click.Extent.Text))
}
$launchRow=[pscustomobject]@{Name='Synthetic-Job';InManifest=$true}
$script:Controls.LaunchJobsGrid.ItemsSource=@($launchRow);$script:Controls.LaunchJobsGrid.SelectedItem=$launchRow
try{
    $window.WindowState='Normal';$window.WindowStartupLocation='Manual';$window.Width=900;$window.Height=600;$window.Left=-20000;$window.Top=-20000;$window.ShowActivated=$false;$window.ShowInTaskbar=$false;$window.Show()
    # 1. Slow success from Operations > Jobs: wait window at once, UI responsive, guarded entries.
    Reset-Case -Delay 1500
    $clock=[Diagnostics.Stopwatch]::StartNew()
    $script:Controls.RequestJobButton.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Button]::ClickEvent))
    $clock.Stop()
    Assert-Case ($clock.ElapsedMilliseconds -lt 1000) "Confirmation path blocked the WPF thread ($($clock.ElapsedMilliseconds) ms)."
    Assert-Case ($script:ConfirmCalls -eq 1 -and $script:RequestSubmitBusy -and $null -ne $script:RequestSubmitContext.Worker) 'Confirmation did not start the background submission.'
    $dialog=$script:RequestSubmitDialog
    Assert-Case ($null -ne $dialog -and $dialog.IsVisible -and $dialog.Owner -eq $window) 'Wait window not shown and owned by the GUI.'
    Assert-Case ($dialog.FindName('SubmitMessageText').Text -ceq ('Submitting request for Synthetic-Job{0} Please wait.' -f [char]0x2026)) 'Wait window text.'
    $bar=$dialog.FindName('SubmitProgressBar')
    Assert-Case ($bar.RenderTransform.HasAnimatedProperties) 'Progress indicator is not animated.'
    Assert-Case (-not $script:Controls.RequestRunButton.IsEnabled -and -not $script:Controls.RequestJobButton.IsEnabled -and -not $script:Controls.RefreshButton.IsEnabled) 'Request and refresh actions stayed enabled.'
    $receiveClock=[Diagnostics.Stopwatch]::StartNew();Receive-GuiJobRunSubmission;$receiveClock.Stop()
    Assert-Case ($script:RequestSubmitBusy -and $receiveClock.ElapsedMilliseconds -lt 150) 'Receiving an active submission blocked the WPF thread.'
    $script:BarFrames=@()
    Wait-Submission -During {
        $script:BarFrames+=Get-BarSnapshot $script:RequestSubmitDialog.FindName('SubmitProgressTrack')
        # Double click on both entries (handlers run even on a disabled button) and the selected-row path.
        $script:Controls.RequestJobButton.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Button]::ClickEvent))
        $script:Controls.PlanningGrid.ItemsSource=@($launchRow);$script:Controls.PlanningGrid.SelectedItem=$launchRow
        $script:Controls.RequestRunButton.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Button]::ClickEvent))
        $script:DoubleResult=Request-SelectedJobRun -JobName 'Synthetic-Job'
        # Auto-refresh and manual requests refresh during the wait.
        Invoke-AutoRefresh;Refresh-RequestsView
        $script:CancelFailure='';try{$script:Controls.CancellationReasonBox.Text='x';Start-GuiPipelineCancellation -All}catch{$script:CancelFailure=$_.Exception.Message}
    } -Later {
        # Later dispatcher tick: the animation clock has advanced without any DoEvents.
        $script:BarFrames+=Get-BarSnapshot $script:RequestSubmitDialog.FindName('SubmitProgressTrack')
        $script:DialogDetailDuring=$script:RequestSubmitDialog.FindName('SubmitDetailText').Text
    }
    Assert-Case ($script:UiTicks -gt 30) "WPF dispatcher stalled during the submission ($($script:UiTicks) ticks)."
    Assert-Case ($script:BarFrames.Count -eq 2 -and $script:BarFrames[0] -ne $script:BarFrames[1]) 'Indeterminate indicator did not animate during the wait.'
    Assert-Case ($script:DialogDetailDuring -match 'Recording the request|Reading the published|Loading orchestrator' -and $script:DialogDetailDuring -match '\(\d+ s\)') 'Wait window does not show progress and elapsed time.'
    Assert-Case ($script:ConfirmCalls -eq 1 -and $null -eq $script:DoubleResult) 'Double click opened a second confirmation or submission.'
    Assert-Case ($script:CancelFailure -match 'run request is being submitted') 'Cancellation allowed during submission.'
    Assert-Case (@(Get-Calls).Count -eq 1) "Expected exactly one submission, got $(@(Get-Calls).Count)."
    Assert-Case ($script:SharedReads -eq 0) 'Requests view read shared data on the WPF thread during the submission.'
    Assert-Case ($script:RefreshCalls -eq 1) "Auto-refresh ran during the submission or not after it ($($script:RefreshCalls))."
    $mainThread=[Threading.Thread]::CurrentThread.ManagedThreadId
    Assert-Case (-not (@(Get-Calls)[0] -match "\|$mainThread$")) 'Submission ran on the WPF thread.'
    Assert-Case (@(Get-Calls)[0] -like 'Synthetic-Job|test|False|*') 'Submission lost job, tenant or dependency choice.'
    Assert-Case (-not $dialog.IsVisible -and $null -eq $script:RequestSubmitDialog) 'Wait window not closed after success.'
    Assert-Case ($script:Results.Count -eq 1 -and -not $script:Results[0].DialogOpen -and -not $script:Results[0].Busy) 'Result shown while the wait window was open.'
    $result=$script:Results[0]
    Assert-Case ($result.Title -eq 'Run requested' -and $result.Icon -eq 'Information' -and $result.Message -match '^Request submitted\.' -and $result.Message -match 'Batch: Batch\d+' -and $result.Message -match 'Jobs: 2' -and $result.Message -match 'Operations > Requests') 'Success confirmation lost BatchId, job count or Requests pointer.'
    Assert-Case ($result.Message -match 'no job has started yet' -and $script:Controls.StatusText.Text -match 'not started yet') 'Success announced a started job.'
    Assert-Case ($script:Controls.RequestRunButton.IsEnabled -and $script:Controls.RequestJobButton.IsEnabled -and $script:Controls.RefreshButton.IsEnabled) 'Controls not restored after success.'
    Assert-Case (@($script:RequestRows).Count -eq 1) 'Requests grid not refreshed from the worker.'
    Assert-Case (-not ((Get-Content $script:GuiLogPath) -match "\|$mainThread\|") -and ((Get-Content $script:GuiLogPath) -match 'Run request recorded')) 'Persistent log missing or written on the WPF thread.'
    Assert-Case ($script:Exceptions.Count -eq 0) "Unexpected GUI exception: $($script:Exceptions -join '; ')"
    # 2. Refusal before anything is written: not recorded, controls restored, no retry.
    Reset-Case -Delay 300 -Fail Before
    [void](Request-SelectedJobRun -JobName 'Synthetic-Job' -IncludeDependencies $true);Wait-Submission
    $result=$script:Results[0]
    Assert-Case ($result.Title -eq 'Run request failed' -and $result.Icon -eq 'Error' -and $result.Message -match 'was not recorded' -and $result.Message -match 'Synthetic refusal before writing') 'Refusal not reported as not recorded.'
    Assert-Case (@(Get-Calls).Count -eq 1 -and @(Get-Calls)[0] -like 'Synthetic-Job|test|True|*') 'Refusal retried or lost the dependency choice.'
    Assert-Case ($null -eq $script:RequestSubmitDialog -and $script:Controls.RequestJobButton.IsEnabled -and $script:Controls.RefreshButton.IsEnabled) 'Controls not restored after an error.'
    # 3. Failure after the request folder appeared: uncertain, never resubmitted.
    Reset-Case -Delay 300 -Fail After
    [void](Request-SelectedJobRun -JobName 'Synthetic-Job');Wait-Submission
    $result=$script:Results[0]
    Assert-Case ($result.Title -eq 'Run request uncertain' -and $result.Icon -eq 'Warning' -and $result.Message -match 'may still have been recorded' -and $result.Message -match 'Batch\d+' -and $result.Message -match 'Nothing is resubmitted automatically') 'Possibly recorded request not reported as uncertain.'
    Assert-Case (@(Get-Calls).Count -eq 1) 'Uncertain request was resubmitted.'
    Assert-Case ($result.Message -notmatch 'not recorded') 'Uncertain result reported as not recorded.'
    # 4. Worker failure before submission (module load): not submitted.
    Reset-Case;$script:ModuleOverride=Join-Path $root 'Missing.psm1'
    [void](Request-SelectedJobRun -JobName 'Synthetic-Job');Wait-Submission
    Assert-Case ($script:Results[0].Message -match 'was not submitted' -and @(Get-Calls).Count -eq 0 -and $script:Results[0].Icon -eq 'Error') 'Pre-submission failure not reported as not submitted.'
    # 5. Published jobs missing: not submitted, no API call.
    Reset-Case -MissingJobs $true
    [void](Request-SelectedJobRun -JobName 'Synthetic-Job');Wait-Submission
    Assert-Case ($script:Results[0].Message -match 'Published jobs configuration not found' -and @(Get-Calls).Count -eq 0) 'Missing published jobs not reported.'
    # 6. Display refresh failure after success stays a success.
    Reset-Case -FailRefresh $true
    [void](Request-SelectedJobRun -JobName 'Synthetic-Job');Wait-Submission
    Assert-Case ($script:Results[0].Title -eq 'Run requested' -and $script:Results[0].Message -match 'could not be refreshed') 'Refresh failure confused with submission failure.'
    # 7. Declined confirmation and running cancellation: nothing starts. An active request no longer blocks: the new one is queued.
    Reset-Case;$script:ConfirmResult=$false
    [void](Request-SelectedJobRun -JobName 'Synthetic-Job')
    Assert-Case ($script:ConfirmCalls -eq 1 -and -not $script:RequestSubmitBusy -and $null -eq $script:RequestSubmitDialog -and @(Get-Calls).Count -eq 0) 'Declining started a submission.'
    Reset-Case -Queue 1;Set-GuiRequestRows -Rows @([pscustomobject]@{BatchId='Active1';Status='Running';Created='';Pending=1;JobRows=@()})
    [void](Request-SelectedJobRun -JobName 'Synthetic-Job');Wait-Submission
    Assert-Case ($script:ConfirmCalls -eq 1 -and $script:LastConfirmMessage -match '1 request\(s\) are active or queued' -and @(Get-Calls).Count -eq 1) 'An active request blocked the submission instead of queueing it.'
    Assert-Case ($script:Results[0].Title -eq 'Run requested' -and $script:Results[0].Message -match 'Request queued' -and $script:Results[0].Message -match 'Queue position: 1, behind Active1') "Queued result not reported: $($script:Results[0].Message)"
    Reset-Case;$script:CancellationBusy=$true
    $failure='';try{[void](Request-SelectedJobRun -JobName 'Synthetic-Job')}catch{$failure=$_.Exception.Message}
    $script:CancellationBusy=$false
    Assert-Case ($failure -match 'cancellation is being checked' -and $script:ConfirmCalls -eq 0) 'Request allowed during a cancellation.'
    # 8b. Real orchestration modules in the worker, synthetic shared folder: the request is
    # recorded with the real API (no resident orchestrator watches this folder, nothing runs).
    $actualRoot=Join-Path $root 'ActualShared'
    $null=New-Item -ItemType Directory -Path (Join-Path $actualRoot 'Config') -Force
    $template=Get-Content -LiteralPath (Join-Path $folder 'Orchestrator-Jobs.json.template') -Raw|ConvertFrom-Json -Depth 100
    $realJob=$template.Jobs[0];$realJob.Enabled=$true
    $realJob|Add-Member -NotePropertyName AssignmentMode -NotePropertyValue 'Legacy' -Force;$realJob|Add-Member -NotePropertyName DependsOn -NotePropertyValue @() -Force
    [IO.File]::WriteAllText((Join-Path $actualRoot 'Config/Orchestrator-Jobs.json.txt'),($template|ConvertTo-Json -Depth 100),[Text.UTF8Encoding]::new($false))
    Reset-Case;$script:SharedDataFolderPath=$actualRoot
    $script:ModuleOverride=@('SmartM365.Orchestrator.Management.psm1','SmartM365.Orchestrator.Pipeline.psm1','SmartM365.Orchestrator.Insights.psm1'|ForEach-Object {(Resolve-Path (Join-Path $folder $_)).Path})
    [void](Request-SelectedJobRun -JobName ([string]$realJob.Name));Wait-Submission
    $script:SharedDataFolderPath=$root
    $realBatches=@(Get-ChildItem -LiteralPath (Join-Path $actualRoot 'PipelineRuns') -Directory)
    Assert-Case ($script:Results[0].Title -eq 'Run requested' -and $realBatches.Count -eq 1 -and $script:Results[0].Message -match [regex]::Escape($realBatches[0].Name)) "Real-provider submission failed: $($script:Results[0].Message)"
    Assert-Case (@($script:RequestRows).Count -eq 1 -and $script:RequestRows[0].BatchId -eq $realBatches[0].Name) 'Real-provider Requests refresh did not show the new batch.'
    # 8. Static contract: no DoEvents, no cancel button, closing protected, both entries share the path.
    $guiText=Get-Content -LiteralPath $guiPath -Raw
    Assert-Case ($guiText -notmatch 'DoEvents') 'DoEvents used.'
    $dialogFunction=$ast.Find({param($node)$node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Show-GuiRequestSubmitDialog'},$true).Extent.Text
    Assert-Case ($dialogFunction -notmatch '<Button') 'Wait window offers a cancel button.'
    $closing=$ast.Find({param($node)$node -is [Management.Automation.Language.InvokeMemberExpressionAst] -and $node.Member.Value -eq 'Add_Closing' -and $node.Expression.Extent.Text -eq '$window'},$true)
    Assert-Case ($closing.Extent.Text -match 'RequestSubmitBusy') 'Window closing not protected during a submission.'
    $autoTick=$ast.Find({param($node)$node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Invoke-AutoRefresh'},$true).Extent.Text
    Assert-Case ($autoTick -match 'RequestSubmitBusy') 'Auto-refresh not guarded.'
    Assert-Case ($guiText -match "RequestSubmitTimer','ConfigurationTimer'" -and $guiText -match "'RequestSubmitContext'") 'Submission timer or worker missing from cleanup.'
    "[{0}] PASS: {1} offline run request submission checks. No job launched, no shared write outside a temporary folder." -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$script:Cases
}
finally{
    $script:SubmitTimer.Stop()
    if($script:RequestSubmitContext.Worker){$null=$script:RequestSubmitContext.Worker.Handle.AsyncWaitHandle.WaitOne(20000);$null=Receive-SmartM365OrchestratorGuiWorker -Context $script:RequestSubmitContext}
    Close-GuiRequestSubmitDialog
    $window.Close()
    $resolved=[IO.Path]::GetFullPath($root)
    if($resolved.StartsWith([IO.Path]::GetFullPath([IO.Path]::GetTempPath()),[StringComparison]::OrdinalIgnoreCase) -and (Split-Path $resolved -Leaf) -like 'SmartM365-GuiRunRequest-Test-*'){
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCANgkLFSNizt7eO
# DUjWDKMT464ag/Tj8nEWJqqBtUtcg6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEII0ZOafbc+qb67toczODcpqKsNeZaNMKHnTvX2Of2vkjMA0GCSqG
# SIb3DQEBAQUABIIBgDVebTO78y1xd3h5WWI08vxcD3nVc5fQSlD8F0zg/WK9xKIL
# ONeLm3U8jfXesIeu4cuvYJiQERlMCi8jww2reUORDdeIetPzoCyeAbN+XK2KTKrd
# kEnfoShhP6GFlEmbrfkk0Psx5gfaic2r1mWht7IzIJSA12PbS9BVwXUVW2FDwpY9
# 9CmOOF2ssqNs4/crueja10UZZFzXYSjm9hA2gBTj6ouS5xLf/FlCMHhquq28jT+R
# XofBRJqYnk0iORtMckZ/fuypw41gUq2Aq/LLYGWI/IvfQrEWgx7YEWnf+qsAjVN0
# D6rOwNhWEyyJ8gPmAkQFXbH1jjjpKhTziQdf3mshnNcXLwjlE/DvRTBF75/Tm4Gt
# wlhG7BEsFQi8l1Vca/ijy+kImv2lAYF8SWf8DSwn4gz1prnRIRdmpxNxYzuloFJI
# bGcdG1WNsa3ShRGYA5gkDJcZEbCXUAiYLPWMy77jpOstM9XBjj3insdlk1RGDf3k
# Inzvqi09DVpyBFNxYqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjgxMDA2
# NDRaMC8GCSqGSIb3DQEJBDEiBCDjCAreeTZjNJnZsX/dZhoJMjSiYpTCTZ9mn0dZ
# EKQB8zANBgkqhkiG9w0BAQEFAASCAgAu4cZ/O0OC7YFuHb8hs9G1apCk2u6mqu1K
# oipvM2CQFpGgSMwhLgFksOnVMGEtqHWjI0cD7VRzA0ckIWf0bc3J3jn2X0xh/+/n
# lgnuS/deqqs+GYCDmABc54H9vDy5ieFQvSa+0KsORViFiW/KgpA7xdjz6XdBOeOL
# fYPeyz7MIq1bcyZWQ+Ha4h67iMAxQN3mKFwG+ocxjGwW+Nr+xDfJ6HHdE6hUWnMC
# ZsK/NCKsGoV2fIKO4OwrcDzdSGIsblIEjrTZImeghBhxbjj8NsRQsxXdsEOBY6rW
# 19k2aR0XqcpeL3XsC4KsYwnX/15giQs4+nbgco268J5C659AnyIO5165I+BldLqI
# BNQmhzVx+w2TmdGhfhzpzxd+jkWypf4IZb5s6gPvGU95lqGyL0xlQTDa9TgAjpzg
# DD/xs4Sv29TECSnni+F8TlIb0gw9vuXJPiIUNxmFAB1oXfWncSOx0+s9+sUXrka4
# QuQHSIB4QQCWDaxRgOMTxaB3M84WSggB2+HB4ZAlHjOGg8kPOSZioh37c8nTBURU
# NT4grDIvU1MNdJp8STojf5+/2KjJdR/YW94OShQahrQ3rfvpYp4YeABE2KOn0apj
# ippau0RLW/ByiOlYM39h96/OdiI8HigAAHBhOmbbu0lhLEAHmVVa+djNUY69QBH0
# xSdRAyyFDQ==
# SIG # End signature block
