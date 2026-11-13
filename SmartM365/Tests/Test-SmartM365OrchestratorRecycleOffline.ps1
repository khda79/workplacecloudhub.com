#Requires -Version 7.0
$ErrorActionPreference='Stop'
$base=Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
$source=Join-Path $base 'SmartM365-Inventory-Orchestrator.ps1'
$ast=[Management.Automation.Language.Parser]::ParseFile($source,[ref]$null,[ref]$null)
$names=@('Save-OrchestratorDeferredSharePointUploads','Invoke-OrchestratorPeriodicSharePointUpload','Send-OrchestratorRuntimeUpdateEmailBounded')
$definitions=foreach($name in $names) { $node=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true); if(-not $node){throw "Missing function $name"}; $node.Extent.Text }
Import-Module (Join-Path $PSScriptRoot '../Modules/SmartM365.Core/SmartM365.Core.psd1') -Force
$core=(Get-Command Send-SmartM365Mail).Module
if($core.Name -ne 'SmartM365.Core'){throw 'Core mail module not available.'}
$runtime=New-Module -ScriptBlock ([scriptblock]::Create($definitions -join "`n"))
$root=Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-RecycleOffline-'+[guid]::NewGuid().ToString('N'))
try {
    $null=New-Item -ItemType Directory -Path $root -Force
    $mail1=Join-Path $root 'one.htm'; $mail2=Join-Path $root 'two.htm'
    [IO.File]::WriteAllText($mail1,'synthetic 1'); [IO.File]::WriteAllText($mail2,'synthetic 2')
    $result=& $runtime {
        param($Root,$Mail1,$Mail2,$Core)
        $script:Settings=@{SendMailMode='Graph';SharedDataFolderPath=$Root;OrchestratorLogFolderPath=$Root;OrchestratorSharePointUploadIntervalMinutes=60;OrchestratorRunsCsvPath='';StatePath='';HeartbeatPath=''}
        $script:State=@{Jobs=@{}}; $script:LastSharePointUploadAttempt=[datetime]::MinValue; $script:Uploaded=[Collections.Generic.List[string]]::new(); $script:Saved=0; $script:SeenDeadline=$null
        $global:SmartM365MailHtmlFiles=@($Mail1,$Mail2)
        function script:Write-OrchestratorLog {param($Message,$Level)}
        function script:Save-OrchestratorState {$script:Saved++}
        function script:Get-OrchestratorLogPath {param($Date) ''}
        function script:Get-JobRunsCsvPath {''}
        function script:Invoke-OrchestratorSharePointUpload {param($LocalFilePath,$Reason,[switch]$Force) $script:Uploaded.Add($LocalFilePath); -not $script:FailUpload}
        function script:Invoke-OrchestratorSharePointMirror {param([switch]$Force)}
        function script:Send-OrchestratorRuntimeUpdateEmail {param($Candidate,$DetectedAt) $script:SeenDeadline=& $Core { $script:SmartM365NetworkDeadline }; $script:ObservedTimeout=& $Core { Get-SmartM365NetworkCallTimeout -TimeoutSeconds 60 }; $true}
        $prior=[pscustomobject]@{ ExpiresAtUtc=[datetime]::UtcNow.AddHours(1);Label='prior';Minutes=60 }
        & $Core {param($value) $script:SmartM365NetworkDeadline=$value} $prior
        $sent=Send-OrchestratorRuntimeUpdateEmailBounded -Candidate ([pscustomobject]@{})
        $restored=& $Core {$script:SmartM365NetworkDeadline}
        $script:Settings.SendMailMode='SMTP'; $smtp=Send-OrchestratorRuntimeUpdateEmailBounded -Candidate ([pscustomobject]@{})
        Save-OrchestratorDeferredSharePointUploads
        Invoke-OrchestratorPeriodicSharePointUpload -Now (Get-Date) -RunNow -SkipMirror
        $during=$script:Uploaded.Count; $queued=@($script:State.DeferredSharePointUploads).Count
        # Simulate a new resident: its in-memory list is empty but saved state survives.
        $global:SmartM365MailHtmlFiles=@()
        Invoke-OrchestratorPeriodicSharePointUpload -Now (Get-Date) -RunNow
        $after=$script:Uploaded.Count; $remaining=@($script:State.DeferredSharePointUploads).Count
        $script:FailUpload=$true
        Invoke-OrchestratorPeriodicSharePointUpload -Now (Get-Date) -RunNow
        $afterFailure=@($script:State.DeferredSharePointUploads).Count
        $script:FailUpload=$false
        Invoke-OrchestratorPeriodicSharePointUpload -Now (Get-Date) -RunNow
        [pscustomobject]@{Sent=$sent;Deadline=$script:SeenDeadline;ObservedTimeout=$script:ObservedTimeout;Restored=$restored;Smtp=$smtp;During=$during;Queued=$queued;After=$after;Remaining=$remaining;AfterFailure=$afterFailure;FinallyRemaining=@($script:State.DeferredSharePointUploads).Count;Saved=$script:Saved}
    } $root $mail1 $mail2 $core
    if(-not $result.Sent -or $null -eq $result.Deadline -or $result.Deadline.Label -ne 'Runtime update notification' -or $result.ObservedTimeout -lt 1 -or $result.ObservedTimeout -gt 15 -or $result.Restored.Label -ne 'prior' -or $result.Smtp){throw 'Bounded Graph notification or deadline restoration failed.'}
    if($result.During -ne 0 -or $result.Queued -ne 2 -or $result.After -ne 1 -or $result.Remaining -ne 1 -or $result.AfterFailure -ne 1 -or $result.FinallyRemaining -ne 0 -or $result.Saved -lt 3){throw 'Deferred uploads were not preserved and retried across recycle.'}
    $finallyNode=@($ast.FindAll({param($n) $n -is [Management.Automation.Language.TryStatementAst] -and $n.Finally -and $n.Finally.Extent.Text -match 'Complete-OrchestratorRunTracking'},$true))[0]
    $cleanup=$finallyNode.Finally.Extent.Text
    if($cleanup.IndexOf('Save-OrchestratorDeferredSharePointUploads') -lt 0 -or $cleanup.IndexOf('Save-OrchestratorDeferredSharePointUploads') -gt $cleanup.IndexOf('Exit-OrchestratorLock') -or $cleanup -notmatch 'SkipMirror:\(Test-OrchestratorPlannedRecycle\)'){throw 'Finalization does not save deferred copies before releasing the resident lock.'}
    'PASS: recycle bounds notification and persists deferred uploads'
} finally {
    & $core {$script:SmartM365NetworkDeadline=$null}
    Remove-Module $runtime -ErrorAction SilentlyContinue
    if(Test-Path -LiteralPath $root){Remove-Item -LiteralPath $root -Recurse -Force}
}
