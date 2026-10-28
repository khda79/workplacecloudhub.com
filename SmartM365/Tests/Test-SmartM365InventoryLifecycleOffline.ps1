<#
.SYNOPSIS
    Offline lifecycle regression: no collector, tenant initialization, preparation or provider request.
.VERSION
    1.0.5
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
$globals=@('SmartM365InventoryRun','EnableSharePointUpload','EnableTeamsNotifications','SmartM365ExecutionSummaryWritten','SmartM365ExecutionStatus','SmartM365ExecutionStartTime','SmartM365WarningCount','SmartM365ErrorCount','SmartM365ScriptName','SmartM365ProfileKey','SmartM365TenantKey','SmartM365OrganizationKey','SmartM365EnvironmentKey','SmartM365TenantId','BasePath','LogTextFile','logTranscriptFile','csvGeneratedPaths','SmartM365SharePointUploadedFiles','SmartM365MailHtmlFiles')
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
    Assert-Lifecycle ($targets.Count -eq 57) 'Approved envelope count changed'
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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB5YcvffeOoUi4H
# Hi1fR5z6AT/3Gbs5bpivNs3i2bpMgKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIBJuIL+oiMTKrUghjtrJRTGFSzGkAokgyIlr8XUfV+qHMA0GCSqG
# SIb3DQEBAQUABIIBgFaBe5Zq1MgygLJTGUF/CPz9O76EGjP3b8b2ImGww2rcmM4j
# RaRXV+5WFvXdYGldz2N7HCLjsk3awDHdpgj5KR/NrduceOUhSEyLEBiesI+29k5g
# VKgiURi3tmgQtyQscavSrm/ZAgpqetw6cFePdWr67mkqz8IaxcLb0fPv08ggoQQo
# stiT1pG1VlUCaBCsOoMh00cx6zdWHYt7EtTXi6o9O3MTaoiYwMwnuR+kuDUinewI
# EqoeCz7W3UVltb/mADn8VHA2WRsKq4BdiSBetM0n7ry7GZJ7HRu8mUezVnrH1KyM
# C86RBPIehK8hONJunW4+zSalOFViWM9kRkc8WEl3JM2VHwWSnerwFWNte/M5i9D2
# 4evcykPmUgJcZgNMt2WbsckIEfISpo6f6le28e+pOevSvg/7smnZhlX4/V5JbJMR
# AYj81xdXKytfQRidRZNr7n1kWYJB+A4eW/OQl1k8BbPcJr/Y6EpnK19gN/Pfh/L8
# yLjL+UIluLhrao8JWKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjQxNDIy
# NTFaMC8GCSqGSIb3DQEJBDEiBCA/c5suXeGwt1GwRIgykx6S6z7w1+EYCFIb0oyZ
# qrkqxjANBgkqhkiG9w0BAQEFAASCAgAyJo81pG+5pNrgMzB+RN7NdKLoI5JTsxXi
# xZHVCiEsXVl0RMSXaeWl15c/gzeOMUsmU9f6TsoQUr+8R2RcHb1Z1MYmmLCy9P3E
# CN44n0kkQHC1mpDhn24/hbdVuS7WhazgtOHhM67zWJArauGRAVROkLkHJShMjexy
# vY9q9FYHS0aVjVDZa42TV9kvUbZYrtwl2ANvTowTu9H49ySe9/F+kAOrSXYHQxK+
# 0vltyP/qtOfZbNFZ6jquiAv9KcUBzSg/6LarSHdQ2zN2pdzjeElTNu2tf+TzMYcW
# 46/HGiEkT8DntNSnAMwcb37cMN3gxaJz/OEgH76Yp8zQBO2X/z2LOSbvP/NgeTA2
# nYGfqrp6oXgtY0Kc5z1CLmOreQ1/B5Ecy50V47iMQyta497WYfRW6EynotWWplQ/
# mYuOBc4VMBg6PTn+DJCgXwWAhR2Pxm8NWMqW6MJl9DJsaY2A1ikcjCfi7upAtP95
# v6Bha60IkBFQ1pu6JPtIxZc6GDNOovwUsr/9dMaeflH7M84+XHBKJ/qA7d0La2RZ
# /JGK/N0mbQPTe82adb2unICzU5d71lkZpY5TG3DnRf8wJBLR8q1ZAnu7Q3WaR7y7
# t0wOIrrEV+xHPJNQidxJHUwCWzYPuaf669cabVdubzPW9+BLnhJm/zwI/N7Wy/ST
# rX1O4nyUFw==
# SIG # End signature block
