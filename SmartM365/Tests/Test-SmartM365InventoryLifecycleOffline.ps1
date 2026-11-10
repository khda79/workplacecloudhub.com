<#
.SYNOPSIS
    Offline lifecycle regression: no collector, tenant initialization, preparation or provider request.
.VERSION
1.0.8
#>
[CmdletBinding()]
param([switch]$ReceiptOnly)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$testRoot=Split-Path $PSScriptRoot -Parent
$helperPath=Join-Path $testRoot 'Modules/SmartM365.Core/SmartM365-InventoryLifecycle.ps1'
. $helperPath
$script:Checks=0
function Assert-Lifecycle([bool]$Condition,[string]$Message){$script:Checks++;if(-not $Condition){throw $Message}}
function Read-LifecycleAst([string]$Path){
    $t=$null;$e=$null;$ast=[Management.Automation.Language.Parser]::ParseFile($Path,[ref]$t,[ref]$e)
    Assert-Lifecycle ($e.Count -eq 0) "Parser: $Path"
    return $ast
}
function Read-LifecycleFunction([string]$Path,[string]$Name){
    $ast=Read-LifecycleAst $Path
    $functions=@($ast.FindAll({param($node)$node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $Name},$true))
    Assert-Lifecycle ($functions.Count -eq 1) "Function not unique: $Name"
    return $functions[0].Extent.Text
}
function Read-LifecycleCleanupText($Finally){
    $text=$Finally.Extent.Text
    Assert-Lifecycle ($text.Contains('if($smartM365RunGuard.ExitCode -eq 130){exit 130}')) 'Cancellation exit belongs to terminal finally'
    # Runspace/process tests exercise the native exit; in-process resource checks omit it.
    $text=$text.Replace('if($smartM365RunGuard.ExitCode -eq 130){exit 130}','')
    return $text.Substring(1,$text.Length-2)
}
$fixture=Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-LifecycleOffline-'+[guid]::NewGuid().ToString('N'))
$globals=@('SmartM365InventoryRun','SmartM365CollectorUploadPolicies','SmartM365CollectorUploadRoots','EnableSharePointUpload','EnableTeamsNotifications','SmartM365ExecutionSummaryWritten','SmartM365ExecutionStatus','SmartM365ExecutionStartTime','SmartM365WarningCount','SmartM365ErrorCount','SmartM365ScriptName','SmartM365ProfileKey','SmartM365TenantKey','SmartM365OrganizationKey','SmartM365EnvironmentKey','SmartM365TenantId','BasePath','LogTextFile','logTranscriptFile','csvGeneratedPaths','SmartM365SharePointUploadedFiles','SmartM365MailHtmlFiles')
$saved=@{};foreach($name in $globals){$v=Get-Variable $name -Scope Global -ErrorAction SilentlyContinue;$saved[$name]=@{Exists=($null-ne$v);Value=$(if($v){$v.Value}else{$null})}}
$previousExit=[Environment]::ExitCode
$script:Events=New-Object 'Collections.Generic.List[string]'
$script:SmartM365SourceReceiptUploadPaths=@()
function WriteLog {param([string]$Message,[string]$Level) if($Level -eq 'WARNING'){$global:SmartM365WarningCount++};if($Level -eq 'ERROR'){$global:SmartM365ErrorCount++}}
function Complete-SmartM365CmdbSourceReceipt {param([string]$Status,[int]$ErrorCount) $script:Events.Add('Receipt:'+ $Status)}
function Invoke-SmartM365SharePointCsvUpload {param([string]$LocalFilePath) if(Test-SmartM365InterruptedRun){throw 'Upload after cancellation'}; $script:Events.Add('Upload')}
function Test-SmartM365SourceReceiptMirrorBlocked {param([string]$Path) $false}
function Write-SmartM365CompletionBanner {
    param([string]$Status,[string]$ScriptName,$StartedAt,$EndedAt,[int]$WarningCount,[int]$ErrorCount,[int]$GeneratedCsvFiles,[string]$LogPath,[string]$ClosedTranscriptPath,[switch]$Force,[switch]$NoConsole,[switch]$ConsoleOnly)
    $status=Get-SmartM365InventoryCompletionStatus -Status $Status -ErrorRecord $null
    if($status -eq 'Auto'){$status='Success'}
    $script:Events.Add('Banner:'+ $Status)
}
function Disconnect-MgGraph {[CmdletBinding()]param() $script:Events.Add('DisconnectGraph')}
function Disconnect-ExchangeOnline {[CmdletBinding()]param([bool]$Confirm) $script:Events.Add('DisconnectExchange')}
function Disconnect-MicrosoftTeams {[CmdletBinding()]param() $script:Events.Add('DisconnectTeams')}
function Disconnect-SmartM365CloudSession {param([bool]$ExchangeOnline,[bool]$Graph,[bool]$VerboseDisconnect) $script:Events.Add('DisconnectCalendar')}
function Reset-LifecycleFixture {
    # Collector entry initialization is omitted when extracting terminal cleanup.
    $script:ExchangeOnPremContext=$null
    foreach($name in $globals){Set-Variable $name -Scope Global -Value $null}
    $global:EnableSharePointUpload=$true;$global:EnableTeamsNotifications=$true
    $global:SmartM365ExecutionSummaryWritten=$false;$global:SmartM365WarningCount=0;$global:SmartM365ErrorCount=0
    $global:SmartM365ExecutionStartTime=Get-Date;$global:SmartM365ScriptName='OfflineLifecycle'
    $global:BasePath=$fixture;$global:LogTextFile=Join-Path $fixture 'trace.log';$global:logTranscriptFile=$null
    $global:csvGeneratedPaths=@();$global:SmartM365SharePointUploadedFiles=@();$global:SmartM365MailHtmlFiles=@()
    $script:Events.Clear();[Environment]::ExitCode=0
}
try {
    $outerPolicy=@{marker='outer-policy'};$outerRoots=@{marker='outer-roots'}
    $global:SmartM365CollectorUploadPolicies=$outerPolicy;$global:SmartM365CollectorUploadRoots=$outerRoots
    $policyState=Start-SmartM365InventoryRun -ScriptPath 'synthetic-policy.ps1'
    Assert-Lifecycle ($null -eq (Get-Variable SmartM365CollectorUploadPolicies -Scope Global -ValueOnly -ErrorAction SilentlyContinue)) 'Nested run inherited an upload policy.'
    Assert-Lifecycle ($null -eq (Get-Variable SmartM365CollectorUploadRoots -Scope Global -ValueOnly -ErrorAction SilentlyContinue)) 'Nested run inherited upload roots.'
    $global:SmartM365CollectorUploadPolicies=@{marker='inner-policy'};$global:SmartM365CollectorUploadRoots=@{marker='inner-roots'}
    $policyState.ReachedTerminal=$true
    Complete-SmartM365InventoryRun -State $policyState
    Assert-Lifecycle ([object]::ReferenceEquals($global:SmartM365CollectorUploadPolicies,$outerPolicy)) 'Outer upload policy was not restored.'
    Assert-Lifecycle ([object]::ReferenceEquals($global:SmartM365CollectorUploadRoots,$outerRoots)) 'Outer upload roots were not restored.'
    # An inner collector must not replace the outer run's effective upload guard.
    foreach($outerEnabled in @($false,$true)) {
        $global:EnableSharePointUpload=$outerEnabled
        $outerRun=Start-SmartM365InventoryRun -ScriptPath 'synthetic-outer.ps1'
        $innerRun=Start-SmartM365InventoryRun -ScriptPath 'synthetic-inner.ps1'
        $global:EnableSharePointUpload=-not $outerEnabled
        $innerRun.ManagedCollectorLegacyUpload=@{Value=(-not $outerEnabled)}
        $innerRun.ReachedTerminal=$true
        Complete-SmartM365InventoryRun -State $innerRun
        Assert-Lifecycle ($global:EnableSharePointUpload -eq $outerEnabled) 'Nested collector did not restore the outer upload guard.'
        $outerRun.ReachedTerminal=$true
        Complete-SmartM365InventoryRun -State $outerRun
    }
    New-Item -ItemType Directory -Path $fixture|Out-Null
    [IO.File]::WriteAllText((Join-Path $fixture 'trace.log'),'synthetic trace')
    $core=Join-Path $testRoot 'Modules/SmartM365.Core/SmartM365.Core.psm1'
    # The broad suite parses PowerShell 7 collectors; receipt-only also runs on 5.1.
    if(-not $ReceiptOnly){
    Invoke-Expression (Read-LifecycleFunction $core 'Complete-SmartM365ExecutionContext')
    foreach($implementation in @($core,(Join-Path $testRoot 'Modules/SmartM365.Core/Compatibility/WindowsPowerShell5/SmartM365-WindowsPowerShell5.psm1'))){
        Invoke-Expression (Read-LifecycleFunction $implementation 'Complete-SmartM365ExecutionContext')
        foreach($scenario in @('Success','Failed','Cancelled')){
            Reset-LifecycleFixture
            $state=Start-SmartM365InventoryRun 'module-completion.ps1' -RequireReturn;$state.Initialized=$true
            if($scenario -ne 'Cancelled'){
                Complete-SmartM365ExecutionContext -Status $scenario
                $state.ReturnedNormally=$true
                Assert-Lifecycle ($script:Events.Count -eq 0) "Module completion deferred: $implementation"
            }
            Complete-SmartM365InventoryRun $state
            Assert-Lifecycle ($script:Events.Contains('Banner:'+ $scenario)) "Module terminal status $scenario : $implementation"
            if($scenario -eq 'Cancelled'){Assert-Lifecycle (-not $script:Events.Contains('Upload') -and $state.ExitCode -eq 130) 'Cancelled module publishes no trace'}
        }
    }
    Invoke-Expression (Read-LifecycleFunction $core 'Complete-SmartM365ExecutionContext')
    $inventory=Join-Path $testRoot 'SmartInventory'
    $targets=@(Get-ChildItem -LiteralPath $inventory -Recurse -File -Filter '*.ps1' | Where-Object {
        $_.FullName -notmatch '[\\/](Tests|TestRuns)[\\/]' -and
        [IO.File]::ReadAllText($_.FullName).Contains('Start-SmartM365InventoryRun -ScriptPath $PSCommandPath')
    })
    Assert-Lifecycle ($targets.Count -eq 59) "Approved envelope count changed: $($targets.Count)"
    foreach($target in $targets){
        $ast=Read-LifecycleAst $target.FullName
        $outer=@($ast.EndBlock.Statements|Where-Object {$_ -is [Management.Automation.Language.TryStatementAst] -and $_.Finally -and $_.Finally.Extent.Text.Contains('Complete-SmartM365InventoryRun -State')})
        Assert-Lifecycle ($outer.Count -eq 1) "Entry envelope: $($target.Name)"
        Assert-Lifecycle ($outer[0].CatchClauses.Count -gt 0) "Outer error preservation: $($target.Name)"
        Assert-Lifecycle ($ast.Extent.Text.Contains('-ScriptPath $PSCommandPath -RequireReturn')) "Normal return required: $($target.Name)"
        Assert-Lifecycle ($outer[0].Body.Extent.Text.Contains('$smartM365RunGuard.ReturnedNormally=$true')) "Terminal return marker: $($target.Name)"
        foreach($node in $ast.FindAll({param($n)$n -is [Management.Automation.Language.ReturnStatementAst]},$true)){
            $nested=$false
            for($parent=$node.Parent;$parent;$parent=$parent.Parent){if($parent -is [Management.Automation.Language.ScriptBlockAst] -and $parent -ne $ast){$nested=$true;break}}
            if($nested){
                $prefix=$ast.Extent.Text.Substring([Math]::Max(0,$node.Extent.StartOffset-100),[Math]::Min(100,$node.Extent.StartOffset))
                Assert-Lifecycle (-not $prefix.EndsWith('$smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; ')) "Helper return is not entry completion: $($target.Name)"
            }
        }
        foreach($scenario in @('Success','Error','Cancelled','CancelledAfterBusiness')){
            Reset-LifecycleFixture
            # Entry variables declared before the envelope try (Windows 11 readiness run lock).
            $runLock=$null
            $smartM365RunGuard=Start-SmartM365InventoryRun -ScriptPath $target.FullName -RequireReturn
            $smartM365RunGuard.Initialized=$true
            $smartM365RunGuard.ReachedTerminal=$scenario -in @('Success','CancelledAfterBusiness')
            $smartM365RunGuard.ReturnedNormally=$scenario -eq 'Success'
            if($scenario -eq 'Error'){$smartM365RunGuard.ErrorRecord=[Management.Automation.ErrorRecord]::new([InvalidOperationException]::new('Synthetic error'),'Offline.Error',[Management.Automation.ErrorCategory]::InvalidOperation,$null)}
            & ([scriptblock]::Create((Read-LifecycleCleanupText $outer[0].Finally)))
            $expected=if($scenario -eq 'Error'){'Failed'}elseif($scenario -like 'Cancelled*'){'Cancelled'}else{$scenario}
            Assert-Lifecycle ($script:Events.Contains('Banner:'+ $expected)) "Terminal status $scenario : $($target.Name)"
            Assert-Lifecycle ($smartM365RunGuard.ExitCode -eq $(if($scenario-like'Cancelled*'){130}elseif($scenario-eq'Error'){1}else{0})) "Exit code $scenario : $($target.Name)"
            Assert-Lifecycle ($global:EnableSharePointUpload -and $global:EnableTeamsNotifications) "Settings restored: $($target.Name)"
            Assert-Lifecycle ($null -eq $global:SmartM365InventoryRun) "Run state released: $($target.Name)"
            Assert-Lifecycle ($smartM365RunGuard.CleanupWarnings.Count -eq 0) "Unexpected cleanup error: $($target.Name)"
            if($scenario -like 'Cancelled*'){Assert-Lifecycle (-not $script:Events.Contains('Upload')) "No upload $($target.Name)"}
            $count=$script:Events.Count;Complete-SmartM365InventoryRun $smartM365RunGuard
            Assert-Lifecycle ($script:Events.Count -eq $count) "Idempotent cleanup: $($target.Name)"
        }
    }
    Reset-LifecycleFixture
    $state=Start-SmartM365InventoryRun -ScriptPath 'synthetic.ps1';$state.Initialized=$true
    Complete-SmartM365ExecutionContext -Status Success
    Assert-Lifecycle ($script:Events.Count -eq 0) 'Completion is deferred until resources close'
    $state.GraphOwned=$true;$state.ExchangeOwned=$true;$state.TeamsOwned=$true
    Complete-SmartM365InventoryRun $state
    Assert-Lifecycle ($script:Events.IndexOf('DisconnectGraph') -lt $script:Events.IndexOf('Banner:Success')) 'Graph closed before banner'
    Assert-Lifecycle ($script:Events.Contains('DisconnectExchange') -and $script:Events.Contains('DisconnectTeams')) 'Owned cloud connections closed'
    Reset-LifecycleFixture
    $parent=Start-SmartM365InventoryRun 'parent.ps1';$parent.ReachedTerminal=$true
    $child=Start-SmartM365InventoryRun 'child.ps1';$child.ReachedTerminal=$true
    Complete-SmartM365InventoryRun $child
    Assert-Lifecycle ([object]::ReferenceEquals($global:SmartM365InventoryRun,$parent)) 'Nested run restores parent identity'
    Complete-SmartM365InventoryRun $parent
    # Exercise the actual entry cleanup callbacks with owned synthetic resources.
    Reset-LifecycleFixture
    $smartM365RunGuard=Start-SmartM365InventoryRun 'ad-cleanup.ps1' -RequireReturn
    $script:workerJobs=@([pscustomobject]@{Id=123;State='Running'})
    $script:tempFolder=$fixture;$script:baseFolder=$fixture
    function Stop-Job {param($Job) $Job.State='Stopped';$script:Events.Add('ADStopped')}
    function Remove-Job {param($Job) $script:Events.Add('ADRemoved')}
    function Remove-TemporaryInventoryFolder {param($TempFolder,$BaseFolder) Assert-Lifecycle ($TempFolder -eq $fixture -and $BaseFolder -eq $fixture) 'AD cleanup ownership';$script:Events.Add('ADTempClosed')}
    $adAst=Read-LifecycleAst (Join-Path $inventory 'ActiveDirectoryInventory/SmartM365-ActiveDirectory-Inventory.ps1')
    $adFinal=@($adAst.EndBlock.Statements|Where-Object {$_ -is [Management.Automation.Language.TryStatementAst] -and $_.Finally -and $_.Finally.Extent.Text.Contains('Complete-SmartM365InventoryRun -State')})[0].Finally
    & ([scriptblock]::Create((Read-LifecycleCleanupText $adFinal)))
    Assert-Lifecycle ($script:Events.Contains('ADStopped') -and $script:Events.Contains('ADRemoved') -and $script:Events.Contains('ADTempClosed')) 'AD owned workers and staging close on interruption'
    Remove-Variable workerJobs,tempFolder,baseFolder -Scope Script
    Reset-LifecycleFixture
    $smartM365RunGuard=Start-SmartM365InventoryRun 'proxy-cleanup.ps1' -RequireReturn
    $script:ChangedViewEntireForest=$true;$script:PrevViewEntireForest=$false
    function Set-ADServerSettings {param([bool]$ViewEntireForest) Assert-Lifecycle (-not $ViewEntireForest) 'Previous AD session setting';$script:Events.Add('ADSettingsRestored')}
    $proxyAst=Read-LifecycleAst (Join-Path $inventory 'ExchangeInventory/OnPremises/ProxyAddresses/SmartM365-Check-ProxyAddresses-Exchange.ps1')
    $proxyFinal=@($proxyAst.EndBlock.Statements|Where-Object {$_ -is [Management.Automation.Language.TryStatementAst] -and $_.Finally -and $_.Finally.Extent.Text.Contains('Complete-SmartM365InventoryRun -State')})[0].Finally
    & ([scriptblock]::Create((Read-LifecycleCleanupText $proxyFinal)))
    Assert-Lifecycle ($script:Events.Contains('ADSettingsRestored') -and -not $script:ChangedViewEntireForest) 'AD session restored on interruption'
    Remove-Variable ChangedViewEntireForest,PrevViewEntireForest -Scope Script
    Reset-LifecycleFixture
    $smartM365RunGuard=Start-SmartM365InventoryRun 'mailbox-cleanup.ps1' -RequireReturn
    $InventoryCompletedSuccessfully=$false
    $mailboxAst=Read-LifecycleAst (Join-Path $inventory 'ExchangeInventory/OnPremises/Mailboxes/SmartM365-Exchange-Local-Mailboxes-Inventory.ps1')
    $mailboxOuter=@($mailboxAst.EndBlock.Statements|Where-Object {$_ -is [Management.Automation.Language.TryStatementAst] -and $_.Finally -and $_.Finally.Extent.Text.Contains('Complete-SmartM365InventoryRun -State')})[0]
    $mailboxMain=@($mailboxOuter.Body.Statements|Where-Object {$_ -is [Management.Automation.Language.TryStatementAst] -and $_.Finally -and $_.Finally.Extent.Text.Contains('if (-not $InventoryCompletedSuccessfully')})[0]
    Assert-Lifecycle ($mailboxMain.CatchClauses.Count -gt 0) 'Mailbox failure preserved before final cleanup'
    $mailboxFinal=$mailboxMain.Finally.Extent.Text
    & ([scriptblock]::Create($mailboxFinal.Substring(1,$mailboxFinal.Length-2)))
    Complete-SmartM365InventoryRun $smartM365RunGuard
    Assert-Lifecycle ($smartM365RunGuard.ExitCode -eq 130 -and -not $smartM365RunGuard.ErrorRecord) 'Mailbox cancellation is not replaced by a generic failure'
    # Local pipeline cancellation during processing and during inner finally.
    foreach($stage in @('Processing','Cleanup','BespokeProcessing','BespokeCleanup')){
        $control=[hashtable]::Synchronized(@{Ready=$false;Status='';Released=$false})
        $actor=[powershell]::Create()
        try {
            $code={
                param($helperPath,$control,$stage)
                . $helperPath
                function Complete-SmartM365ExecutionContext {param($Status,$ErrorRecord,$FailureStage) $control.Status=Get-SmartM365InventoryCompletionStatus $Status $ErrorRecord}
                $bespoke=$stage -like 'Bespoke*';$stage=$stage -replace '^Bespoke',''
                $state=Start-SmartM365InventoryRun 'synthetic-entry.ps1' -RequireReturn;$state.Initialized=-not $bespoke
                if($bespoke){$state.PendingCompletion=@{Status='Success'}}
                try {
                    try {
                        if($stage -eq 'Processing'){$control.Ready=$true;while($true){[Threading.Thread]::Sleep(10)}}
                        $state.ReachedTerminal=$true
                    } finally {
                        $state.InCleanup=$true
                        if($stage -eq 'Cleanup'){$control.Ready=$true;while($true){[Threading.Thread]::Sleep(10)}}
                    }
                    $state.ReturnedNormally=$true
                } finally {
                    Complete-SmartM365InventoryRun $state
                    $control.Released=($state.Finalized -and $null -eq (Get-Variable SmartM365InventoryRun -Scope Global -ValueOnly -ErrorAction SilentlyContinue))
                    $control.ExitCode=$state.ExitCode
                }
            }
            $null=$actor.AddScript($code.ToString()).AddArgument($helperPath).AddArgument($control).AddArgument($stage)
            $handle=$actor.BeginInvoke();$deadline=[datetime]::UtcNow.AddSeconds(10)
            while(-not $control.Ready -and [datetime]::UtcNow -lt $deadline){[Threading.Thread]::Sleep(10)}
            Assert-Lifecycle ([bool]$control.Ready) "Synthetic cancellation stage reached: $stage"
            $actor.Stop()
            try{$null=$actor.EndInvoke($handle)}catch{if($_.Exception.InnerException -isnot [Management.Automation.PipelineStoppedException]){throw}}
            Assert-Lifecycle ($control.Status -eq 'Cancelled' -and $control.ExitCode -eq 130 -and $control.Released) "Pipeline cancellation finalized: $stage"
        } finally {$actor.Dispose()}
    }
    # Native hosts must receive the code from the actual entry-point finally,
    # not only from Environment.ExitCode (which hosts can replace with zero).
    $nativeAst=Read-LifecycleAst (Join-Path $inventory 'M365Inventory/WorkplaceScope/SmartM365-WorkplaceScope-Inventory.ps1')
    $nativeFinally=@($nativeAst.EndBlock.Statements|Where-Object {$_ -is [Management.Automation.Language.TryStatementAst] -and $_.Finally -and $_.Finally.Extent.Text.Contains('Complete-SmartM365InventoryRun -State')})[0].Finally.Extent.Text
    $nativeCode=@'
param([string]$Scenario,[string]$Helper)
. $Helper
$smartM365RunGuard=Start-SmartM365InventoryRun 'synthetic-native.ps1' -RequireReturn
try {
    try {
        if($Scenario -eq 'Failed'){throw 'Synthetic terminal error'}
        if($Scenario -eq 'Cancelled'){throw [Management.Automation.PipelineStoppedException]::new()}
        $smartM365RunGuard.ReachedTerminal=$true
    } finally {
        $smartM365RunGuard.InCleanup=$true
        if($Scenario -eq 'CancelledAfterBusiness'){throw [Management.Automation.PipelineStoppedException]::new()}
    }
    $smartM365RunGuard.ReturnedNormally=$true
} catch {
    if($_.Exception -isnot [Management.Automation.PipelineStoppedException]){$smartM365RunGuard.ErrorRecord=$_}
    throw
} finally
'@
    $nativePath=Join-Path $fixture 'native-exit.ps1'
    [IO.File]::WriteAllText($nativePath,$nativeCode+$nativeFinally,[Text.UTF8Encoding]::new($true))
    foreach($engineName in @('pwsh.exe','powershell.exe')){
        $engine=(Get-Command $engineName -ErrorAction Stop).Source
        foreach($scenario in @('Success','Failed','Cancelled','CancelledAfterBusiness')){
            $child=New-Object Diagnostics.Process
            try {
                $child.StartInfo.FileName=$engine
                $child.StartInfo.Arguments='-NoProfile -ExecutionPolicy Bypass -File "'+$nativePath+'" -Scenario '+$scenario+' -Helper "'+$helperPath+'"'
                $child.StartInfo.UseShellExecute=$false;$child.StartInfo.CreateNoWindow=$true
                $child.StartInfo.RedirectStandardOutput=$true;$child.StartInfo.RedirectStandardError=$true
                $null=$child.Start()
                if(-not $child.WaitForExit(10000)){throw 'Synthetic native lifecycle process exceeded its timeout'}
                $expected=if($scenario -eq 'Failed'){1}elseif($scenario -like 'Cancelled*'){130}else{0}
                Assert-Lifecycle ($child.ExitCode -eq $expected) "Native $engineName code $expected : $scenario; received $($child.ExitCode)"
            } finally {
                # This is only the process just created with our unique fixture path.
                if($child.Id -and -not $child.HasExited){$child.Kill();$null=$child.WaitForExit(1000)}
                $child.Dispose()
            }
        }
    }
    Reset-LifecycleFixture
    $state=Start-SmartM365InventoryRun 'cleanup.ps1';$state.Initialized=$true;$state.ReachedTerminal=$true
    Invoke-SmartM365RunCleanup $state 'synthetic failure' {throw 'Synthetic cleanup error'}
    Invoke-SmartM365RunCleanup $state 'remaining resource' {$script:Events.Add('RemainingClosed')}
    Complete-SmartM365InventoryRun $state
    Assert-Lifecycle ($state.CleanupWarnings.Count -eq 1 -and $script:Events.Contains('RemainingClosed')) 'Cleanup errors do not skip remaining resources'
    Assert-Lifecycle ($script:Events.Contains('Banner:CompletedWithWarnings')) 'Cleanup failure is not reported as clean success'
    Reset-LifecycleFixture
    $state=Start-SmartM365InventoryRun 'owned-transcript.ps1' -RequireReturn
    $tracePath=Join-Path $fixture 'owned-transcript.txt'
    Start-SmartM365RunTranscript -Path $tracePath | Out-Null
    Assert-Lifecycle ([bool]$state.TranscriptOwned) 'Native transcript ownership recorded'
    Complete-SmartM365InventoryRun $state
    Assert-Lifecycle (-not $state.TranscriptOwned) 'Native transcript stopped on cancellation'
    $traceHandle=[IO.File]::Open($tracePath,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
    $traceHandle.Dispose()
    Assert-Lifecycle ($state.CleanupWarnings.Count -eq 0) 'Native transcript handle released'
    foreach($modulePath in @($core,(Join-Path $testRoot 'Modules/SmartM365.Core/Compatibility/WindowsPowerShell5/SmartM365-WindowsPowerShell5.psm1'))){
        foreach($name in @('Invoke-SmartM365SharePointCsvUpload','SendEmailHtmlReport','Send-SmartM365TeamsNotification','RemoveOldFiles','Remove-SmartM365TimestampedFilesOlderThan')){
            $code=Read-LifecycleFunction $modulePath $name
            Assert-Lifecycle ($code.Contains('if (Test-SmartM365InterruptedRun) { return }')) "No external/retention action after interruption: $name"
        }
    }
    $workerModule=Join-Path $inventory 'Orchestrator/SmartM365.Orchestrator.GuiWorker.psm1'
    Import-Module $workerModule -Force
    $context=[pscustomobject]@{Worker=$null;Progress=[hashtable]::Synchronized(@{})}
    Start-SmartM365OrchestratorGuiWorker -Context $context -InputObject @{} -Work {
        param($progress,$inputObject)
        try{$progress.Ready=$true;while($true){[Threading.Thread]::Sleep(10)}}finally{$progress.Closed=$true}
    }|Out-Null
    $deadline=[datetime]::UtcNow.AddSeconds(5)
    while(-not $context.Progress.ContainsKey('Ready') -and [datetime]::UtcNow -lt $deadline){[Threading.Thread]::Sleep(10)}
    Close-SmartM365OrchestratorGuiWorker -Context $context
    Assert-Lifecycle ($null -eq $context.Worker -and $context.Progress.Closed) 'GUI worker cancellation releases its runspace and runs finally'
    Close-SmartM365OrchestratorGuiWorker -Context $context
    $detector=Join-Path $inventory 'M365Inventory/IntuneInventory/Devices/SmartM365-Detect-DeviceSystemInfo.ps1'
    $detectionCode=[IO.File]::ReadAllText($detector)
    Assert-Lifecycle ($detectionCode.Contains('FinalReleaseComObject($updateSystemInfo)')) 'COM release is present'
    $consoleCode=Read-LifecycleFunction $detector 'Get-LastConsoleLogon'
    Assert-Lifecycle ($consoleCode.Contains('$securityEvent.Dispose()')) 'Event records released, including early return'
    Assert-Lifecycle ($detectionCode.Contains("'DeviceSystemPayloadEnd:1.6'")) 'Detection payload contract unchanged'
    }
    # Real receipt finalization with bespoke initialization (Autopatch's path).
    # Only the shared definition helpers run; all transport/log/banner calls remain mocks.
    . (Join-Path $testRoot 'Modules/SmartM365.Core/SmartM365-CmdbReceipt.ps1')
    function Test-SmartM365MaxItemsMode { return $false }
    function Get-SmartM365ScriptVersionFromFile { param($Path) $null=$Path; return 'synthetic-test' }
    $producer='SmartM365-Intune-WindowsAutopatch-Alerts-Inventory.ps1'
    $definition=(Get-Content $script:SmartM365CmdbRegistryPath -Raw | ConvertFrom-Json).Producers | Where-Object Script -eq $producer
    foreach($implementation in @($core,(Join-Path $testRoot 'Modules/SmartM365.Core/Compatibility/WindowsPowerShell5/SmartM365-WindowsPowerShell5.psm1'))){
        foreach($prefix in @('','Core')){
            Remove-Item Function:Complete-SmartM365ExecutionContext -ErrorAction SilentlyContinue
            Remove-Item Function:Complete-CoreSmartM365ExecutionContext -ErrorAction SilentlyContinue
            $command='Complete-'+$prefix+'SmartM365ExecutionContext'
            $completionCode=Read-LifecycleFunction $implementation 'Complete-SmartM365ExecutionContext'
            Invoke-Expression ($completionCode.Replace('function Complete-SmartM365ExecutionContext',('function '+$command)))
            $receiptFolder=Join-Path $fixture ([guid]::NewGuid().ToString('N'))
            foreach($scenario in @('Success','Failed','Cancelled','MissingOutput')){
                Reset-LifecycleFixture
                $global:SmartM365TenantKey='synthetic';$global:SmartM365OrganizationKey='test'
                $global:SmartM365EnvironmentKey='test';$global:SmartM365TenantId='synthetic-tenant'
                $state=Start-SmartM365InventoryRun $producer -RequireReturn
                Start-SmartM365CmdbSourceReceipt -ScriptPath $producer -SourceRootPath $receiptFolder
                $receiptPath=$script:SmartM365CmdbSourceContext.ReceiptPath
                $runPath=$script:SmartM365CmdbSourceContext.RunPath
                $proofBefore=if(Test-Path $receiptPath){(Get-FileHash $receiptPath).Hash}else{''}
                $csvNames=@('Intune_WindowsAutopatch_Alerts_PolicySummary.csv','Intune_WindowsAutopatch_Alerts_Detail.csv','Intune_WindowsAutopatch_Alerts_Summary.csv')
                foreach($csvName in $csvNames){
                    if($scenario -eq 'MissingOutput' -and $csvName -in $definition.Files){continue}
                    $csvPath=Join-Path $receiptFolder $csvName
                    Start-SmartM365SourcePublication -Path $csvPath
                    [IO.File]::WriteAllText($csvPath,("TenantKey,Name`r`nsynthetic,"+$scenario+"`r`n"))
                    Register-SmartM365SourceCsv -Path $csvPath
                }
                $csvHashes=@(Get-ChildItem $receiptFolder -Filter '*.csv' | Get-FileHash | Select-Object Path,Hash) | ConvertTo-Json -Compress
                $state.InCleanup=$true;$state.ReachedTerminal=$scenario -ne 'Cancelled'
                Set-SmartM365CmdbSourceScope -CompleteScope $true -Scope $definition.Scope
                $requested=if($scenario -eq 'Failed'){'Failed'}else{'Success'}
                & $command -Status $requested
                Assert-Lifecycle ($null -ne $state.PendingCompletion -and -not $state.Initialized) 'Bespoke completion queued without standard initialization'
                Assert-Lifecycle ((Get-Content $runPath -Raw | ConvertFrom-Json).Status -eq 'Publishing') 'Receipt finalized before outer cleanup'
                $state.ReturnedNormally=$scenario -ne 'Cancelled'
                Complete-SmartM365InventoryRun $state
                $run=Get-Content $runPath -Raw | ConvertFrom-Json
                $expected=if($scenario -eq 'Success'){'Completed'}else{'Failed'}
                Assert-Lifecycle ($run.Status -eq $expected -and $run.CompletedAtUtc) "Receipt terminal state: $scenario / $prefix"
                Assert-Lifecycle ($run.UnqualifiedPublication -eq ($scenario -ne 'Success')) 'Incomplete run cleared publication fence'
                Assert-Lifecycle ($state.CleanupWarnings.Count -eq 0 -and $null -eq $script:SmartM365CmdbSourceContext) 'Receipt finalization failed or lock context retained'
                Assert-Lifecycle ((@(Get-ChildItem $receiptFolder -Filter '*.csv' | Get-FileHash | Select-Object Path,Hash) | ConvertTo-Json -Compress) -ceq $csvHashes) 'Finalization changed CSV content'
                if($scenario -eq 'Success'){
                    $proof=Get-Content $receiptPath -Raw | ConvertFrom-Json
                    Assert-Lifecycle ($proof.Status -eq 'Completed' -and $proof.RunId -eq $run.RunId -and $proof.Files.Count -eq 3) 'Current proof does not describe completed run'
                }else{
                    Assert-Lifecycle ((Get-FileHash $receiptPath).Hash -eq $proofBefore) 'Failed attempt overwrote last completion proof'
                }
                if($scenario -eq 'Failed'){Assert-Lifecycle ($state.ExitCode -eq 1) 'Bespoke failure returned zero'}
                if($scenario -eq 'Cancelled'){Assert-Lifecycle ($state.ExitCode -eq 130 -and -not $script:Events.Contains('Upload')) 'Cancelled bespoke run published or returned success'}
                $terminalHash=(Get-FileHash $runPath).Hash
                Complete-SmartM365InventoryRun $state
                Assert-Lifecycle ((Get-FileHash $runPath).Hash -eq $terminalHash) 'Receipt finalized more than once'
            }
        }
    }
    # Early initialization failure without an explicit completion must not invent one.
    Reset-LifecycleFixture
    $state=Start-SmartM365InventoryRun 'not-initialized.ps1' -RequireReturn
    $state.ErrorRecord=[Management.Automation.ErrorRecord]::new([InvalidOperationException]::new('Synthetic initialization error'),'Offline.Init',[Management.Automation.ErrorCategory]::InvalidOperation,$null)
    Complete-SmartM365InventoryRun $state
    Assert-Lifecycle ($script:Events.Count -eq 0 -and $state.ExitCode -eq 1) 'Uninitialized run fabricated a terminal summary'
    "PASS: $script:Checks lifecycle assertions. No collection, preparation, transfer or Power BI refresh."
}
finally {
    if($null -ne (Get-Variable SmartM365CmdbSourceContext -Scope Script -ValueOnly -ErrorAction SilentlyContinue) -and $script:SmartM365CmdbSourceContext.Lock){$script:SmartM365CmdbSourceContext.Lock.Dispose()}
    [Environment]::ExitCode=$previousExit
    foreach($name in $saved.Keys){if($saved[$name].Exists){Set-Variable $name -Scope Global -Value $saved[$name].Value}else{Remove-Variable $name -Scope Global -ErrorAction SilentlyContinue}}
    if($null -ne (Get-Variable context -ValueOnly -ErrorAction SilentlyContinue) -and $context.Worker){try{Close-SmartM365OrchestratorGuiWorker $context}catch{}}
    $resolved=[IO.Path]::GetFullPath($fixture)
    if((Split-Path $resolved -Leaf) -match '^SmartM365-LifecycleOffline-[a-f0-9]{32}$' -and (Split-Path $resolved -Parent).TrimEnd('\') -eq ([IO.Path]::GetTempPath()).TrimEnd('\')){
        if(Test-Path -LiteralPath $resolved){Remove-Item -LiteralPath $resolved -Recurse -Force}
    }
}

# SIG # Begin signature block
# MIIfPQYJKoZIhvcNAQcCoIIfLjCCHyoCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDrChaesmx8tStL
# NDEuXBRD63peYyPhMAySY8FY97hFdaCCGNgwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBoIw
# ggRqoAMCAQICEDbCsL18Gzrno7PdNsvJdWgwDQYJKoZIhvcNAQEMBQAwgYgxCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpOZXcgSmVyc2V5MRQwEgYDVQQHEwtKZXJzZXkg
# Q2l0eTEeMBwGA1UEChMVVGhlIFVTRVJUUlVTVCBOZXR3b3JrMS4wLAYDVQQDEyVV
# U0VSVHJ1c3QgUlNBIENlcnRpZmljYXRpb24gQXV0aG9yaXR5MB4XDTIxMDMyMjAw
# MDAwMFoXDTM4MDExODIzNTk1OVowVzELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1Nl
# Y3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFt
# cGluZyBSb290IFI0NjCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAIid
# 2LlFZ50d3ei5JoGaVFTAfEkFm8xaFQ/ZlBBEtEFAgXcUmanU5HYsyAhTXiDQkiUv
# pVdYqZ1uYoZEMgtHES1l1Cc6HaqZzEbOOp6YiTx63ywTon434aXVydmhx7Dx4IBr
# Aou7hNGsKioIBPy5GMN7KmgYmuu4f92sKKjbxqohUSfjk1mJlAjthgF7Hjx4vvyV
# DQGsd5KarLW5d73E3ThobSkob2SL48LpUR/O627pDchxll+bTSv1gASn/hp6IuHJ
# orEu6EopoB1CNFp/+HpTXeNARXUmdRMKbnXWflq+/g36NJXB35ZvxQw6zid61qmr
# lD/IbKJA6COw/8lFSPQwBP1ityZdwuCysCKZ9ZjczMqbUcLFyq6KdOpuzVDR3ZUw
# xDKL1wCAxgL2Mpz7eZbrb/JWXiOcNzDpQsmwGQ6Stw8tTCqPumhLRPb7YkzM8/6N
# nWH3T9ClmcGSF22LEyJYNWCHrQqYubNeKolzqUbCqhSqmr/UdUeb49zYHr7ALL8b
# AJyPDmubNqMtuaobKASBqP84uhqcRY/pjnYd+V5/dcu9ieERjiRKKsxCG1t6tG9o
# j7liwPddXEcYGOUiWLm742st50jGwTzxbMpepmOP1mLnJskvZaN5e45NuzAHteOR
# lsSuDt5t4BBRCJL+5EZnnw0ezntk9R8QJyAkL6/bAgMBAAGjggEWMIIBEjAfBgNV
# HSMEGDAWgBRTeb9aqitKz1SA4dibwJ3ysgNmyzAdBgNVHQ4EFgQU9ndq3T/9ARP/
# FqFsggIv0Ao9FCUwDgYDVR0PAQH/BAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wEwYD
# VR0lBAwwCgYIKwYBBQUHAwgwEQYDVR0gBAowCDAGBgRVHSAAMFAGA1UdHwRJMEcw
# RaBDoEGGP2h0dHA6Ly9jcmwudXNlcnRydXN0LmNvbS9VU0VSVHJ1c3RSU0FDZXJ0
# aWZpY2F0aW9uQXV0aG9yaXR5LmNybDA1BggrBgEFBQcBAQQpMCcwJQYIKwYBBQUH
# MAGGGWh0dHA6Ly9vY3NwLnVzZXJ0cnVzdC5jb20wDQYJKoZIhvcNAQEMBQADggIB
# AA6+ZUHtaES45aHF1BGH5Lc7JYzrftrIF5Ht2PFDxKKFOct/awAEWgHQMVHol9ZL
# Syd/pYMbaC0IZ+XBW9xhdkkmUV/KbUOiL7g98M/yzRyqUOZ1/IY7Ay0YbMniIibJ
# rPcgFp73WDnRDKtVutShPSZQZAdtFwXnuiWl8eFARK3PmLqEm9UsVX+55DbVIz33
# Mbhba0HUTEYv3yJ1fwKGxPBsP/MgTECimh7eXomvMm0/GPxX2uhwCcs/YLxDnBdV
# VlxvDjHjO1cuwbOpkiJGHmLXXVNbsdXUC2xBrq9fLrfe8IBsA4hopwsCj8hTuwKX
# JlSTrZcPRVSccP5i9U28gZ7OMzoJGlxZ5384OKm0r568Mo9TYrqzKeKZgFo0fj2/
# 0iHbj55hc20jfxvK3mQi+H7xpbzxZOFGm/yVQkpo+ffv5gdhp+hv1GDsvJOtJinJ
# mgGbBFZIThbqI+MHvAmMmkfb3fTxmSkop2mSJL1Y2x/955S29Gu0gSJIkc3z30vU
# /iXrMpWx2tS7UVfVP+5tKuzGtgkP7d/doqDrLF1u6Ci3TpjAZdeLLlRQZm867eVe
# XED58LXd1Dk6UvaAhvmWYXoiLz4JA5gPBcz7J311uahxCweNxE+xxxR3kT0WKzAS
# o5G/PyDez6NHdIUKBeE3jDPs2ACc6CkJ1Sji4PKWVT0/MIIGpzCCBI+gAwIBAgIR
# AJCsCHIg/cWnxGtcxw33PQYwDQYJKoZIhvcNAQEMBQAwVzELMAkGA1UEBhMCR0Ix
# GDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJs
# aWMgVGltZSBTdGFtcGluZyBSb290IFI0NjAeFw0yNjAzMjUwMDAwMDBaFw00MTAz
# MjQyMzU5NTlaMFUxCzAJBgNVBAYTAkdCMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0
# ZWQxLDAqBgNVBAMTI1NlY3RpZ28gUHVibGljIFRpbWUgU3RhbXBpbmcgQ0EgUjQx
# MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAruRKogGtghxi+WYtW5oD
# zPDVF8GGSfgbUKh6bONxi0wvrI1S8qbAYfvLr/ky5ILVRg//70pgNKq8xC3/WEQo
# djEwAP2hmkGShoNAUQps4kd6Wwp74Fo7RlwQ1Mp949ytpWQDvCsYBbZccDmBAJC/
# ggqiuL/c805fGcMw6TzIgyBWuUx5PGp9YnheSNPXFzaz0MPtREdZYk4WhtM+hazq
# asMWVpj0WUAcNhN9vO/FAdWy9Gafdb7lmYLDKTTYjwqAY9P9RfixPPjUaJH6mnBS
# NBdrX7a0Qdlux0ApS0fc48RW1m+W3tq3HiHzch1FHyhiLzCNjc6MUpcV5xalBvPO
# w/FtQo/AxaJOvPCSsVrx0f/WkMpEm3fvVbrY9+oo9rIKv9ducE6VGfwIAtKYedG0
# bO4Ba1MmlxPcErDqjLwggvrBJu73fwXpkhtE0hzV0psgm2vhQs3pHll9N00SHBdy
# 2qndEcNuDh+46XouM2hoXCO533YQQOHPEUnMTWOo3hyxx5kjDE5PVqp+x+HS4VAT
# +WBMG4GzeLr9YvZbU5x5YvLdcR1dErV/QRYK55rp019fZFF2NR+TkSW0WcmQ3b5t
# aGcrXg49EpzKM6/mEpnSJXg1E13X6GO29rWs/LNvkGzsS8XGoRCGBls6ruofeebS
# sHADR3GeIE5gIU927bjokLECAwEAAaOCAW4wggFqMB8GA1UdIwQYMBaAFPZ3at0/
# /QET/xahbIICL9AKPRQlMB0GA1UdDgQWBBQ6dKUMZ8ZCUML9tfzHuyk0gvR6uTAO
# BgNVHQ8BAf8EBAMCAYYwEgYDVR0TAQH/BAgwBgEB/wIBADATBgNVHSUEDDAKBggr
# BgEFBQcDCDAjBgNVHSAEHDAaMAgGBmeBDAEEAjAOBgwrBgEEAbIxAQIBAwgwTAYD
# VR0fBEUwQzBBoD+gPYY7aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVi
# bGljVGltZVN0YW1waW5nUm9vdFI0Ni5jcmwwfAYIKwYBBQUHAQEEcDBuMEcGCCsG
# AQUFBzAChjtodHRwOi8vY3J0LnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWNUaW1l
# U3RhbXBpbmdSb290UjQ2LnA3YzAjBggrBgEFBQcwAYYXaHR0cDovL29jc3Auc2Vj
# dGlnby5jb20wDQYJKoZIhvcNAQEMBQADggIBADLeUkdm8Z4DZvjfKHOhqu+hsdXN
# t+X5F+48PB6PTAJCRgA3qxxO3YbAV7baps4K/+2WWoWUspBkT4T2NXK49NvJAfCs
# ztHSgtqkAQMLX4KjCypJF/+m2Ktrk993g+gcgKdv1yg9C3JmYdJCnnL0ga+pZ/Wo
# 1+rtXZ8dnwO8RCstTN6gYX0ElFi7Y7NpxbdBC1S6bc05V/SA9HC/ojj33W6Gdwnp
# U/iVylSkdkoHtHeGIhQLT2ZH0qPM9Wdce8v2fZsDCJQQJ8rll7OGLDbsXa2CLf0M
# RN9TwzifQ3rEuAXOx/TkzkZRFfwL34hf1XqSmaYq2tTMy2LgsPrqC2Z/6ZKb3fgr
# zU0vphB4wSTWulitY/KlxbvoyKvrBvUCCx4sgeqf8aR65CbvM5MN/d/lahfXipU2
# NlY0cXcnGS61XpmeGKd8It92/lufApZR9x6o5qMJWe0jq4JsfGMGDpIKx7FzkB8g
# aejuBUW/CJ9Phc40+xJRonvVewn4S9yJVRWeM47irGbR9YlN3xruM/yZzhk+rAm9
# AW06nv7ob6RQkAXR+cTxiAPy620FF41NrViYB4UyKpzfx7x8jh4ubTOMz954YIdq
# yeiqqtsbBwXjWLP0dfMUPA3iIPnPdBKGnodGJTdSlPAMmKJdyvTPqmOXs/LMnf+2
# Za0Z6FXsIB9z9aXLMIIG4jCCBMqgAwIBAgIRAOdO8lWwUE/626bf9/yLoxUwDQYJ
# KoZIhvcNAQEMBQAwVTELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBDQSBS
# NDEwHhcNMjYwMzI1MDAwMDAwWhcNMzcwNjI0MjM1OTU5WjByMQswCQYDVQQGEwJH
# QjEXMBUGA1UECBMOR3JlYXRlciBMb25kb24xGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEwMC4GA1UEAxMnU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBTaWdu
# ZXIgUjM3MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAsv/DbUvcUNlF
# LQURd9m4+1St5+JudFKo5P803Iks4mFeNB9SymodP6BJJWBuNhOFQj9w77AVAeg5
# qQpA2dIwp2QTyBHr2h9eWSTkMBVj9mV6+WI5SaW+vDZW7PhJTbysd9v9WB3Xt6ql
# Ei8m47pcTy8+k/OfhziKiuzNQXqfC7KcoRD/6up8OZBsU0qxr7n5nh/iRfAp1QXF
# TBQONBZSGIdHAyVRYYX033VoC8v71rizEKCpH97Pxbwcn9eq9K7W8h5v4npsMUoq
# CS/c8mQwylDQGx15dHYV6NlcVFdjXD11l7qCrIy/unH5OlZtgx58QJRXRbGgQyBd
# STpEpwuj3i5Qc52Z9m7hd7yCGCXKujf83hUQpOPx1w8+84EbEUTHVAfq4cpORaGW
# gY8NJy6txmd3wpS1MeXrOaVAMczTgzAZ+yZBWIqdgQBgTxEeXldEToZOrRkxvn1I
# jIlfr4I4NWJz+Rb52FshLVnkA/wdoad789Eb7XZDNKd4oMmnc636TgauaaVZP2LL
# oU0JD/fYr53hwBn4uXu5ZsSfpnqAT60S7szJm/Na882xEoyRzLJ+UVbXOlHLO63D
# KkAtdz1CDuwWxgRE1drnwplepT06dz+1yTr5p1AkUz21bzE6cT/8/kjh4OPzggYY
# qrOBQPfuKEL5ZJPcN9jRgEpYvRlq5ucCAwEAAaOCAY4wggGKMB8GA1UdIwQYMBaA
# FDp0pQxnxkJQwv21/Me7KTSC9Hq5MB0GA1UdDgQWBBRhEOl6Eq9RxIXU8s+kdA9Q
# zSCv+DAOBgNVHQ8BAf8EBAMCBsAwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAK
# BggrBgEFBQcDCDBKBgNVHSAEQzBBMAgGBmeBDAEEAjA1BgwrBgEEAbIxAQIBAwgw
# JTAjBggrBgEFBQcCARYXaHR0cHM6Ly9zZWN0aWdvLmNvbS9DUFMwSgYDVR0fBEMw
# QTA/oD2gO4Y5aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGlt
# ZVN0YW1waW5nQ0FSNDEuY3JsMHoGCCsGAQUFBwEBBG4wbDBFBggrBgEFBQcwAoY5
# aHR0cDovL2NydC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGltZVN0YW1waW5n
# Q0FSNDEuY3J0MCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5zZWN0aWdvLmNvbTAN
# BgkqhkiG9w0BAQwFAAOCAgEAA+o9jdGszfoZepOmygef1OlbkjrPd2QW9z3M8vVb
# QSCruPeO2eRsC9GhZ4CMZfhkrixayYD67gQkbyiRCbJu5L/i0NQjlQhBvbWfiEba
# +KHFKGud5YHRWhDZUtDeMIJGZG0BD7/sftZUo2Ifk+CXi/ZlM50+xK3OkqeXVi5G
# ubDD/5txmYuqCT3T3LAilmoB+5th9sQxiMhyQuT3R/aYb4vypoZJLYklUzTalXle
# W1nV9s4UROlE389CHDKAi/fepRSMnV8TghODDQxwzNGrOJZ04k/yhzHHDupfHPU5
# 1FYJqXIvWq9SAAWdlNV1JGIxhkp/TAtxBwz/Vd/VbgVb2d9/wRFfxFkka39O0+4x
# aZSl/oEK/1DqjxjJRO2Se9lGlJDScu21Zd23Cys3aYyB8y5H/+DFWtVe8PMKgr+V
# uIDp0Rk5bneVDAEW0TPAT8Ufwl2F6DJiDg/KZk5NmsYES+CxvF7bnISEnQh0ZrWn
# AJixquV0mElUx01wA5TuPIgyodxzNq/fC0hen9LBtdnfFfSZ+wt8A1Injsbio+DH
# Vq1voYiVNpBfO7+nh9NB4AhRXNldPgr3zgjJ+47s0uNYy2iDXAZSlkP3ym/7gy31
# jlu989SNpRWO14/LUNV2LSuXkRI1iLTPI6ZdXG0DnPPG7UftF0tk5m6BP9eNfr2t
# j1sxggW7MIIFtwIBATBiME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNv
# bTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91ZGh1Yi5jb20C
# EB5u7zvHOOGzRDjI8QMf4mgwDQYJYIZIAWUDBAIBBQCggYQwGAYKKwYBBAGCNwIB
# DDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEE
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgribxYFCKSkme
# 77wjR0lE7YX4OFl/vcsMPlhJ5z/BI7kwDQYJKoZIhvcNAQEBBQAEggGAZgy+pBoM
# RCYCBAhogkAQgU2at/9qjq5qTRct/QVc5J8jhfiHxgq7wDbi/EOnhDOHmcVjLRQg
# ZnY1F41AnJSjJ6gjyd5mwnyD8dpXIAFeFcawpeS+i2vQXh9eUOTCB22bryAUfF2Z
# 5+UFfTc6IBMTRz6wgSm6Q7K9+hql942v/b5qIBxL/gSchCNX6Z5BvcDaFBVcgeoZ
# Gm3M5AJ52WABNfrkrmQzPLB1V6+kl9Ofi1wMspuDM56+sDO999NABMazBGlQ5km8
# EDUMurSZe4wUbq5GyTig7bPJIu4hablVjp8C/AFRGE+0fXn28U+OwKD0BxJiApP6
# j5yQNuomqvhfFPxb6dO0Qnz9XFfZg32Rsh5YLTEu4SF1wCFtyzz20hXPWjAPoFso
# O9VwvDBejIdxBP/v2YuNsBOISnylhWEBKiiVIIaBHCTc05SnRES3/3czNwScdyEk
# uQniv3Z0PNpOuuUmWS6984rtm2Roe6LtkUGY8KpBYa6a0S8ssWh2yoYBoYIDIzCC
# Ax8GCSqGSIb3DQEJBjGCAxAwggMMAgEBMGowVTELMAkGA1UEBhMCR0IxGDAWBgNV
# BAoTD1NlY3RpZ28gTGltaXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGlt
# ZSBTdGFtcGluZyBDQSBSNDECEQDnTvJVsFBP+tum3/f8i6MVMA0GCWCGSAFlAwQC
# AgUAoHkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcN
# MjYxMTAxMTQzMjM2WjA/BgkqhkiG9w0BCQQxMgQwx6gpdVjXve6du/OYri+XOKky
# AeSY0PzGrr/LuMKp2vf2V+cKjNXV46sU+mlP8cqLMA0GCSqGSIb3DQEBAQUABIIC
# ACL42vZbAGxaM73D1p79EjGn39c9x5zfauTNzFwBzWvtM2SymdO45rgvKZ1w0U72
# OzsiL+qf1gHMzfupPN97ArGIXEplVviIIq0dAtUJV/WzTlcipUKcuDrbyNcszQSH
# bussW5U45a5xaJXCD0WfdAl4vJR38GoEPRWxQiwYIk8wV6IyeeDnw0E4fhpOLMia
# y1NGpXsVxAldVdAaVtgnzQ30tMxKjxb5h4iI72v2lVRhLTUG5snDcr29UvsboI1W
# 5hLk7034fT9Z1E33VvXxYMw8rLe2LmyuqbWbLzCr5iumGjfGJyXHlRxtoweb2epn
# 4T1XtbLIRxlwr+52vnaVaDq0Kr5dpxQU1jKXizcV3yMBlKgzQ8NJVxcply5j1O1G
# tzPobLauZdD8en5dvRFXOXyeOyFvaIvwKub02ArX6D9BEkMq4sE3oQC3+RRcZBam
# x7YZ4M1JooDkttUWdgz7S6jOGOZeotScG+c8X7NHKnNGhddsIhfHpAe7JOxf822+
# +W9I2MchIom9uieeVOAqDdSiBa80ROZxv8LzEAOgp7b+H5i6zQ1F0eebj0486luk
# jHmWV0gtNbzsviAFPr6eXPKfsDSh/CyplbkHwFUe/vJfCKpn3SzVi8K63J17+kIj
# JvJKZFGi1dTLpoSj0kPpxzjsRGYKHx+p0/6Y7zEt/HpP
# SIG # End signature block
