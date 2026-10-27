#Requires -Version 7.0
<#
.SYNOPSIS
Offline WPF tests of the asynchronous run request submission: wait window, slow, success,
error, uncertain result, double click and auto-refresh during the wait. No job is launched.
.VERSION
1.0.0
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
    [pscustomobject]@{BatchId=$batch;TotalCount=2}
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
function Confirm-GuiJobRunRequest {param($Message) $script:ConfirmCalls++;$script:ConfirmResult}
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
    param([int]$Delay=0,[string]$Fail='',[bool]$FailRefresh=$false,[bool]$MissingJobs=$false)
    Save-State @{DelayMs=$Delay;Fail=$Fail;FailRefresh=$FailRefresh;MissingJobs=$MissingJobs}
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
    Assert-Case ($result.Title -eq 'Run requested' -and $result.Icon -eq 'Information' -and $result.Message -match '^Request submitted\.' -and $result.Message -match 'Batch: Batch\d+' -and $result.Message -match 'Jobs: 2' -and $result.Message -match 'Requests tab') 'Success confirmation lost BatchId, job count or Requests pointer.'
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
    # 7. Declined confirmation, active request and running cancellation: nothing starts.
    Reset-Case;$script:ConfirmResult=$false
    [void](Request-SelectedJobRun -JobName 'Synthetic-Job')
    Assert-Case ($script:ConfirmCalls -eq 1 -and -not $script:RequestSubmitBusy -and $null -eq $script:RequestSubmitDialog -and @(Get-Calls).Count -eq 0) 'Declining started a submission.'
    Reset-Case;Set-GuiRequestRows -Rows @([pscustomobject]@{BatchId='Active1';Status='Running';Created='';Pending=1;JobRows=@()})
    $failure='';try{[void](Request-SelectedJobRun -JobName 'Synthetic-Job')}catch{$failure=$_.Exception.Message}
    Assert-Case ($failure -like '*active pipeline request already exists*' -and $script:ConfirmCalls -eq 0 -and -not $script:RequestSubmitBusy) 'Active request pre-check bypassed.'
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
