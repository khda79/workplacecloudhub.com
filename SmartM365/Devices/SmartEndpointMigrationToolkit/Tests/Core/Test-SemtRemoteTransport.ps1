#requires -Version 5.1
<#
.SYNOPSIS
Synthetic Task Scheduler tests. No scheduled task or remote computer is touched.
.VERSION
1.0.0
#>
$ErrorActionPreference = 'Stop'
class SemtFakeAction { [string]$Path; [string]$Arguments }
class SemtFakeActions {
    [int]$Count = 1
    [SemtFakeAction]$Action = [SemtFakeAction]::new()
    [object] Create([int]$type) { return $this.Action }
    [object] Item([int]$index) { return $this.Action }
}
class SemtFakePrincipal { [string]$UserId; [int]$LogonType; [int]$RunLevel }
class SemtFakeSettings { [bool]$AllowDemandStart; [bool]$DisallowStartIfOnBatteries; [bool]$StopIfGoingOnBatteries; [string]$ExecutionTimeLimit }
class SemtFakeRegistration { [string]$Description }
class SemtFakeDefinition {
    [SemtFakeActions]$Actions = [SemtFakeActions]::new()
    [SemtFakePrincipal]$Principal = [SemtFakePrincipal]::new()
    [SemtFakeSettings]$Settings = [SemtFakeSettings]::new()
    [SemtFakeRegistration]$RegistrationInfo = [SemtFakeRegistration]::new()
}
class SemtFakeTask {
    [SemtFakeDefinition]$Definition
    [int]$State = 3
    [datetime]$LastRunTime = [datetime]::Now
    [int]$LastTaskResult = 0
    [object] Run([object]$arg) { return $null }
}
class SemtFakeFolder {
    [SemtFakeTask]$Task
    [bool]$Deleted = $false
    [int]$Flags
    [bool]$Collide = $false
    [bool]$ChangeAction = $false
    [string]$ReplacementPrincipal = ''
    # Synthetic RPC failures: COMException with the HRESULT seen on real targets.
    [int]$GetTaskFailures = 0
    [bool]$AlwaysFail = $false
    [int]$FailHResult = -2147023170
    [int]$GetTaskCalls = 0
    [bool]$Gone = $false
    [int]$DeleteFailures = 0
    [bool]$DeleteAppliesBeforeFailure = $false
    [object] RegisterTaskDefinition([string]$name,[object]$definition,[int]$flags,[object]$user,[object]$password,[int]$logon,[object]$sddl) {
        if ($this.Collide) { throw 'task already exists' }
        $this.Flags = $flags
        $this.Task = [SemtFakeTask]::new(); $this.Task.Definition = $definition
        return $this.Task
    }
    [object] GetTask([string]$name) {
        $this.GetTaskCalls++
        if ($this.Gone) { throw [System.Runtime.InteropServices.COMException]::new('The system cannot find the file specified. (Exception from HRESULT: 0x80070002)', -2147024894) }
        if ($this.AlwaysFail -or $this.GetTaskFailures -gt 0) {
            $this.GetTaskFailures--
            throw [System.Runtime.InteropServices.COMException]::new(('Synthetic RPC failure. (Exception from HRESULT: 0x{0:X8})' -f $this.FailHResult), $this.FailHResult)
        }
        if ($this.ChangeAction) { $this.Task.Definition.Actions.Action.Arguments = 'changed' }; if ($this.ReplacementPrincipal) { $this.Task.Definition.Principal.UserId = $this.ReplacementPrincipal }; return $this.Task
    }
    [void] DeleteTask([string]$name,[int]$flags) {
        if ($this.DeleteFailures -gt 0) {
            $this.DeleteFailures--
            if ($this.DeleteAppliesBeforeFailure) { $this.Deleted = $true; $this.Gone = $true }
            throw [System.Runtime.InteropServices.COMException]::new('Synthetic RPC failure. (Exception from HRESULT: 0x800706BE)', -2147023170)
        }
        $this.Deleted = $true
    }
}
class SemtFakeService {
    [SemtFakeFolder]$Folder = [SemtFakeFolder]::new()
    [int]$Connects = 0
    [void] Connect() { $this.Connects++ }
    [void] Connect([string]$target) { $this.Connects++ }
    [object] GetFolder([string]$path) { return $this.Folder }
    [object] NewTask([int]$flags) { return [SemtFakeDefinition]::new() }
}
$root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$remote = Join-Path $root 'Modules\SmartEndpointMigration.Core\Functions\Semt.Remote.ps1'
$tokens=$null; $errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile($remote,[ref]$tokens,[ref]$errors)
if ($errors.Count) { throw $errors[0].Message }
$fn=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Invoke-SemtSystemTaskEndpoint'},$true)
$assignment=$fn.Find({param($n) $n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$operation'},$true)
$expression=$assignment.Right.Extent.Text
$operation=[scriptblock]::Create($expression.Substring(1,$expression.Length-2))
function New-Object { param([string]$ComObject) if ($ComObject -ne 'Schedule.Service') { throw 'Unexpected COM object' }; return $script:fake }
# Synthetic clock: every sleep of the operation advances it, so RPC failure windows elapse instantly.
$script:FakeNow=[datetime]'2026-01-01T10:00:00'
function Start-Sleep { param([int]$Seconds) $script:FakeNow=$script:FakeNow.AddSeconds($Seconds) }
function Get-Date { return $script:FakeNow }
$passed=0
foreach ($mode in @('ScheduledTask','WinRM')) {
    $script:fake=[SemtFakeService]::new()
    $dir=Join-Path ([IO.Path]::GetTempPath()) ('SEMT-transport-test-' + [guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($dir)
    $file=Join-Path $dir 'endpoint.ps1'
    [IO.File]::WriteAllText($file,'synthetic')
    $hash=(Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash
    $request=@{Target='synthetic.invalid'; Name='WCH-SEMT-11111111111111111111111111111111'; Arguments='-File "C:\Program Files\test.ps1"'; Script=$file; ShareScript=$file; Sha256=$hash; Mode=$mode}
    $result=& $operation $request
    if (-not $result.Cleaned -or -not $script:fake.Folder.Deleted -or (Test-Path -LiteralPath $dir)) { throw "$mode cleanup failed" }
    if ($script:fake.Folder.Flags -ne 2 -or $script:fake.Folder.Task.Definition.Principal.UserId -ne 'SYSTEM' -or $script:fake.Folder.Task.Definition.Principal.LogonType -ne 5) { throw "$mode SYSTEM or collision guard failed" }
    $passed++
}
$systemSid = [System.Security.Principal.SecurityIdentifier]::new('S-1-5-18')
$localizedName = $systemSid.Translate([System.Security.Principal.NTAccount]).Value
$shortName = ($localizedName -split '\\')[-1]
foreach ($principal in @('S-1-5-18',$localizedName,$shortName)) {
    $script:fake=[SemtFakeService]::new()
    $script:fake.Folder.ReplacementPrincipal=$principal
    $dir=Join-Path ([IO.Path]::GetTempPath()) ('SEMT-transport-test-' + [guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($dir)
    $file=Join-Path $dir 'endpoint.ps1'; [IO.File]::WriteAllText($file,'synthetic')
    $hash=(Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash
    $result=& $operation @{Target='synthetic.invalid';Name='WCH-SEMT-11111111111111111111111111111111';Arguments='test';Script=$file;ShareScript=$file;Sha256=$hash;Mode='ScheduledTask'}
    if (-not $result.Cleaned -or -not $script:fake.Folder.Deleted -or (Test-Path -LiteralPath $dir)) { throw "SYSTEM principal '$principal' was not accepted" }
    $passed++
}
$script:fake=[SemtFakeService]::new()
$script:fake.Folder.ReplacementPrincipal='S-1-5-19'
try { & $operation @{Target='x';Name='x';Arguments='x';Mode='ScheduledTask'}; throw 'Expected rejection of non-SYSTEM principal' }
catch { if ($_.Exception.Message -notmatch 'identity changed') { throw } }
if ($script:fake.Folder.Deleted) { throw 'Non-SYSTEM task was deleted' }
$passed++
$script:fake=[SemtFakeService]::new()
$script:fake.Folder.Collide=$true
try { & $operation @{Target='x';Name='x';Arguments='x';Mode='ScheduledTask'}; throw 'Expected a collision' }
catch { if ($_.Exception.Message -notmatch 'already exists') { throw } }
if ($script:fake.Folder.Deleted) { throw 'Collision deleted a task' }
$passed++
$script:fake=[SemtFakeService]::new()
$script:fake.Folder.ChangeAction=$true
try { & $operation @{Target='x';Name='x';Arguments='x';Mode='ScheduledTask'}; throw 'Expected an identity mismatch' }
catch { if ($_.Exception.Message -notmatch 'identity changed') { throw } }
if ($script:fake.Folder.Deleted) { throw 'Changed task was deleted' }
$passed++
$script:fake=[SemtFakeService]::new()
$dir=Join-Path ([IO.Path]::GetTempPath()) ('SEMT-transport-test-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($dir)
$file=Join-Path $dir 'endpoint.ps1'; [IO.File]::WriteAllText($file,'synthetic')
try {
    try { & $operation @{Target='x';Name='x';Arguments='x';Mode='ScheduledTask';Script=$file;ShareScript=$file;Sha256='wrong'}; throw 'Expected a hash mismatch' }
    catch { if ($_.Exception.Message -notmatch 'hash changed') { throw } }
    if (-not (Test-Path -LiteralPath $file)) { throw 'Changed endpoint was deleted' }
    if ($script:fake.Folder.Deleted) { throw 'Task deleted despite changed endpoint' }
    $passed++
} finally { Remove-Item -LiteralPath $file -Force -ErrorAction SilentlyContinue; Remove-Item -LiteralPath $dir -Force -ErrorAction SilentlyContinue }

# Transient RPC failures while following or cleaning the task.
function Invoke-SemtRpcScenario {
    param([Parameter(Mandatory=$true)]$Service,[int]$Window = 15)
    $script:fake=$Service
    $dir=Join-Path ([IO.Path]::GetTempPath()) ('SEMT-transport-test-' + [guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($dir)
    $file=Join-Path $dir 'endpoint.ps1'; [IO.File]::WriteAllText($file,'synthetic')
    $hash=(Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash
    $startedAt=$script:FakeNow
    $outcome=[ordered]@{ Output=$null; Warnings=@(); Error=''; DirectoryKept=$false; ElapsedMinutes=0 }
    try {
        $records=@(& $operation @{Target='synthetic.invalid';Name='WCH-SEMT-11111111111111111111111111111111';Arguments='test';Script=$file;ShareScript=$file;Sha256=$hash;Mode='ScheduledTask';RpcFailureWindowMinutes=$Window} 3>&1)
        $outcome.Warnings=@($records | Where-Object { $_ -is [System.Management.Automation.WarningRecord] } | ForEach-Object { [string]$_.Message })
        $outcome.Output=@($records | Where-Object { $_ -isnot [System.Management.Automation.WarningRecord] })
    }
    catch { $outcome.Error=$_.Exception.Message }
    finally {
        $outcome.DirectoryKept=(Test-Path -LiteralPath $dir)
        $outcome.ElapsedMinutes=($script:FakeNow - $startedAt).TotalMinutes
        Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue
    }
    return [pscustomobject]$outcome
}
$service=[SemtFakeService]::new(); $service.Folder.GetTaskFailures=3
$scenario=Invoke-SemtRpcScenario -Service $service
if ($scenario.Error -or $scenario.Output.Count -ne 1 -or -not $scenario.Output[0].Cleaned -or -not $service.Folder.Deleted -or $scenario.DirectoryKept) { throw "Intermittent RPC failure was not recovered: $($scenario.Error)" }
if (@($scenario.Warnings | Where-Object { $_ -like 'SEMT_RPC_TRANSIENT*' }).Count -ne 3 -or $scenario.Warnings[2] -notmatch 'ConsecutiveFailures=3; TotalFailures=3' -or $scenario.Warnings[0] -notmatch 'Code=0x800706BE' -or @($scenario.Warnings | Where-Object { $_ -like 'SEMT_RPC_RECOVERED*' }).Count -ne 1) { throw 'RPC failures were not reported with a counter' }
if ($scenario.Output[0].RpcTransientFailures -ne 3 -or $service.Connects -lt 4) { throw 'RPC recovery did not reconnect to the Task Scheduler' }
$passed++
foreach ($code in @(-2147023174,-2147023169,-2147023436,-2147024775,-2147417848)) {
    # 0x800706BA server unavailable, 0x800706BF call failed DNE, 0x800705B4 timeout, 0x80070079 semaphore timeout, 0x80010108 disconnected
    $service=[SemtFakeService]::new(); $service.Folder.GetTaskFailures=2; $service.Folder.FailHResult=$code
    $scenario=Invoke-SemtRpcScenario -Service $service
    if ($scenario.Error -or -not $scenario.Output[0].Cleaned -or -not $service.Folder.Deleted) { throw ('Transient code 0x{0:X8} was not retried: {1}' -f $code,$scenario.Error) }
    $passed++
}
$service=[SemtFakeService]::new(); $service.Folder.AlwaysFail=$true
$scenario=Invoke-SemtRpcScenario -Service $service -Window 15
if ($scenario.Error -notmatch '^SEMT_RPC_UNAVAILABLE: Step=GetTaskState; Code=0x800706BE' -or $scenario.Error -notmatch 'WindowMinutes=15' -or $service.Folder.Deleted -or -not $scenario.DirectoryKept) { throw "Continuous RPC failure was not reported unavailable, or cleanup ran: $($scenario.Error)" }
if ($scenario.ElapsedMinutes -lt 15 -or $scenario.ElapsedMinutes -gt 17) { throw "Failure window not honoured: $($scenario.ElapsedMinutes) minute(s)" }
if ($service.Folder.GetTaskCalls -lt 5 -or $service.Folder.GetTaskCalls -gt 40) { throw "Unexpected retry count: $($service.Folder.GetTaskCalls)" }
$passed++
$service=[SemtFakeService]::new(); $service.Folder.AlwaysFail=$true
$scenario=Invoke-SemtRpcScenario -Service $service -Window 0
if ($scenario.Error -notmatch '^SEMT_RPC_UNAVAILABLE: .*ConsecutiveFailures=1;' -or $service.Folder.GetTaskCalls -ne 1) { throw 'Window 0 did not make the first RPC failure final' }
$passed++
$service=[SemtFakeService]::new(); $service.Folder.GetTaskFailures=1; $service.Folder.FailHResult=-2147024891
$scenario=Invoke-SemtRpcScenario -Service $service
if (-not $scenario.Error -or $scenario.Error -like '*SEMT_RPC_UNAVAILABLE*' -or $scenario.Error -notmatch '0x80070005' -or $service.Folder.GetTaskCalls -ne 1 -or $service.Folder.Deleted) { throw 'A non-transient error (access denied) was retried' }
$passed++
$service=[SemtFakeService]::new(); $service.Folder.DeleteFailures=1
$scenario=Invoke-SemtRpcScenario -Service $service
if ($scenario.Error -or -not $scenario.Output[0].Cleaned -or -not $service.Folder.Deleted -or $scenario.DirectoryKept) { throw "Transient RPC failure on DeleteTask was not retried: $($scenario.Error)" }
$passed++
$service=[SemtFakeService]::new(); $service.Folder.DeleteFailures=1; $service.Folder.DeleteAppliesBeforeFailure=$true
$scenario=Invoke-SemtRpcScenario -Service $service
if ($scenario.Error -or -not $scenario.Output[0].Cleaned -or $scenario.DirectoryKept) { throw "A delete applied before a lost reply was not accepted: $($scenario.Error)" }
$passed++
Remove-Item Function:\New-Object
Remove-Item Function:\Start-Sleep
Remove-Item Function:\Get-Date
Import-Module (Join-Path $root 'Modules\SmartEndpointMigration.Core\SmartEndpointMigration.Core.psd1') -MinimumVersion 1.0.0 -Force
$catalog=Import-SemtOptionCatalog -PackageRoot $root
foreach ($campaign in @('HybridJoin','Windows11Upgrade')) {
    $resolved=Resolve-SemtEffectiveOptions -Catalog $catalog -Campaign $campaign
    if ($resolved['RemoteExecutionMode'].Value -ne 'ScheduledTask') { throw "$campaign default transport differs" }
    $passed++
    foreach ($choice in @('ScheduledTask','WinRM','PsExec')) {
        $resolved=Resolve-SemtEffectiveOptions -Catalog $catalog -Campaign $campaign -CliLayer @{RemoteExecutionMode=$choice}
        if ($resolved['RemoteExecutionMode'].Value -ne $choice) { throw "$campaign transport choice $choice was not retained" }
        $passed++
    }
    try { [void](Resolve-SemtEffectiveOptions -Catalog $catalog -Campaign $campaign -CliLayer @{RemoteExecutionMode='Unexpected'}); throw 'Invalid choice accepted' }
    catch { if ($_.Exception.Message -eq 'Invalid choice accepted') { throw } }
    $passed++
}
$module=Get-Module SmartEndpointMigration.Core
& $module {
    function script:Start-Job { param($ArgumentList,$ScriptBlock); $script:CapturedRequest=$ArgumentList[0]; return [pscustomobject]@{State='Completed'} }
    function script:Wait-Job { param($Job,$Timeout); return $Job }
    function script:Receive-Job { param($Job); return [pscustomobject]@{Cleaned=$true; ExitCode='0'} }
    function script:Remove-Job { param($Job,[switch]$Force) }
}
$log=Join-Path ([IO.Path]::GetTempPath()) ('SEMT-transport-' + [guid]::NewGuid().ToString('N') + '.log')
try {
    foreach ($mode in @('ScheduledTask','WinRM')) {
        $result=Invoke-SemtSystemTaskEndpoint -Mode $mode -ConnectionTarget 'synthetic.invalid' -CorrelationId '11111111111111111111111111111111' -RemoteScriptPath 'C:\Program Files\test.ps1' -RemoteSharePath '\\synthetic.invalid\C$\test.ps1' -ExpectedSha256 'test' -Arguments @('-AuditOnly','-Value','has space') -LogPath $log
        $captured=& $module { $script:CapturedRequest }
        if (-not $result.Completed -or $captured.Mode -ne $mode -or $captured.Arguments -notmatch '"C:\\Program Files\\test.ps1"' -or $captured.Arguments -notmatch '"has space"') { throw "$mode dispatch failed" }
        $passed++
    }
    # Worker side: RPC warnings of the job reach the worker log with a counter, the attempt result
    # is re-read on recovery, and a failure window that elapsed is reported as RpcUnavailable.
    $shareRoot=Join-Path ([IO.Path]::GetTempPath()) ('SEMT-transport-share-' + [guid]::NewGuid().ToString('N'))
    $resultFile=Join-Path $shareRoot 'State\Results\11111111111111111111111111111111.json.txt'
    [void][IO.Directory]::CreateDirectory((Split-Path $resultFile -Parent))
    & $module {
        param($messages)
        $script:FakeWarnings=@($messages | ForEach-Object { [System.Management.Automation.WarningRecord]::new([string]$_) })
        $script:FakeReceive={ [pscustomobject]@{Cleaned=$true; ExitCode='3'; RpcTransientFailures=2} }
        function script:Start-Job { param($ArgumentList,$ScriptBlock); $script:CapturedRequest=$ArgumentList[0]; return [pscustomobject]@{State='Completed'; ChildJobs=@([pscustomobject]@{Warning=$script:FakeWarnings})} }
        function script:Receive-Job { param($Job); & $script:FakeReceive }
    } @('SEMT_RPC_TRANSIENT: Step=GetTaskState; Code=0x800706BE; ConsecutiveFailures=1','SEMT_RPC_TRANSIENT: Step=GetTaskState; Code=0x800706BE; ConsecutiveFailures=2','SEMT_RPC_RECOVERED: Step=GetTaskState; ConsecutiveFailures=2; TotalFailures=2')
    try {
        [IO.File]::WriteAllText($resultFile,'{"CorrelationId":"11111111111111111111111111111111","State":"Completed","Status":"INTUNE_ENROLLMENT_PENDING_CONFIRMATION","ExitCode":"3"}')
        $recoveredTransport=Invoke-SemtSystemTaskEndpoint -Mode ScheduledTask -ConnectionTarget 'synthetic.invalid' -CorrelationId '11111111111111111111111111111111' -RemoteScriptPath 'C:\Program Files\test.ps1' -RemoteSharePath '\\synthetic.invalid\C$\test.ps1' -ExpectedSha256 'test' -Arguments @() -LogPath $log -RpcFailureWindowMinutes 20 -RemoteDataShare $shareRoot
        $captured=& $module { $script:CapturedRequest }
        $logText=[IO.File]::ReadAllText($log)
        if (-not $recoveredTransport.Completed -or $recoveredTransport.RpcTransientFailures -ne 2 -or $recoveredTransport.RpcUnavailable -or $captured.RpcFailureWindowMinutes -ne 20) { throw 'Recovered RPC transport was not reported completed' }
        if ($logText -notmatch 'RPC failure #1 \(retrying' -or $logText -notmatch 'RPC failure #2 \(retrying' -or $logText -notmatch 'recovered after 2 RPC failure' -or $logText -notmatch 'Attempt result after RPC recovery: Attempt result State=Completed') { throw 'RPC failures or recovery were not logged' }
        $recoveredAttempt=Read-SemtRemoteAttemptResult -RemoteDataShare $shareRoot -CorrelationId '11111111111111111111111111111111'
        $passed++
        & $module {
            $script:FakeWarnings=@()
            $script:FakeReceive={ throw 'SEMT_RPC_UNAVAILABLE: Step=GetTaskState; Code=0x800706BE; ConsecutiveFailures=12; TotalFailures=12; FailingForMinutes=15.2; WindowMinutes=15; Error=synthetic' }
        }
        [IO.File]::WriteAllText($resultFile,'{"CorrelationId":"11111111111111111111111111111111","State":"Running"}')
        $lostTransport=Invoke-SemtSystemTaskEndpoint -Mode ScheduledTask -ConnectionTarget 'synthetic.invalid' -CorrelationId '11111111111111111111111111111111' -RemoteScriptPath 'C:\Program Files\test.ps1' -RemoteSharePath '\\synthetic.invalid\C$\test.ps1' -ExpectedSha256 'test' -Arguments @() -LogPath $log -RpcFailureWindowMinutes 15 -RemoteDataShare $shareRoot
        if ($lostTransport.Completed -or -not $lostTransport.RpcUnavailable -or $lostTransport.Detail -notmatch 'RemoteTaskRpcFailureWindowMinutes' -or $lostTransport.Detail -notmatch 'was kept') { throw 'Elapsed RPC window was not reported unavailable' }
        $lostAttempt=Read-SemtRemoteAttemptResult -RemoteDataShare $shareRoot -CorrelationId '11111111111111111111111111111111'
        $passed++
    } finally { Remove-Item -LiteralPath $shareRoot -Recurse -Force -ErrorAction SilentlyContinue }
    & $module {
        function script:Start-Job { param($ArgumentList,$ScriptBlock); $script:CapturedRequest=$ArgumentList[0]; return [pscustomobject]@{State='Completed'} }
        function script:Receive-Job { param($Job); return [pscustomobject]@{Cleaned=$true; ExitCode='0'} }
    }
    & $module {
        $script:FakeClock = 0
        function script:Wait-Job { param($Job,$Timeout); return $null }
        function script:Get-Date { param($Format); $script:FakeClock++; $value=([datetime]'2026-01-01').AddMinutes($script:FakeClock * 2); if ($Format) { return $value.ToString($Format) }; return $value }
        function script:Update-SemtWorkerLease { param($LeasePath) }
    }
    $timeout=Invoke-SemtSystemTaskEndpoint -Mode ScheduledTask -ConnectionTarget 'synthetic.invalid' -CorrelationId '11111111111111111111111111111111' -RemoteScriptPath 'C:\Program Files\test.ps1' -RemoteSharePath '\\synthetic.invalid\C$\test.ps1' -ExpectedSha256 'test' -Arguments @() -LogPath $log -TimeoutMinutes 1
    if (-not $timeout.TimedOut -or $timeout.Completed -or $timeout.Detail -notmatch 'verify before retrying') { throw 'Timeout did not leave the remote state unconfirmed' }
    $passed++
} finally { Remove-Item -LiteralPath $log -Force -ErrorAction SilentlyContinue }
$worker = Join-Path $root 'Scripts\SmartM365-EndpointMigration-PsExecWorker.ps1'
$tokens=$null; $errors=$null
$workerAst=[Management.Automation.Language.Parser]::ParseFile($worker,[ref]$tokens,[ref]$errors)
if ($errors.Count) { throw $errors[0].Message }
$statusBranch=$workerAst.Find({param($n) $n -is [Management.Automation.Language.IfStatementAst] -and $n.Extent.Text.StartsWith('if ($attemptResult.Completed)') -and $n.Extent.Text.Contains('REMOTE_EXECUTION_ACCESS_DENIED')},$true)
if (-not $statusBranch) { throw 'Transport status branch not found' }
$classify=[scriptblock]::Create($statusBranch.Extent.Text)
function Set-WorkerStatus { param($Status,$Detail,$NextAction); $script:observedStatus=$Status; $script:observedNextAction=$NextAction }
foreach ($exitCode in @('-1073741790','3221225506')) {
    $attemptResult=[pscustomobject]@{Completed=$false;Detail='no result'}
    $transport=[pscustomobject]@{Completed=$true;ExitCode=$exitCode;Detail=''}
    $result=[ordered]@{ExitCode='';JobErrorMessage=''}
    $remoteMode='ScheduledTask'
    & $classify
    if ($script:observedStatus -ne 'REMOTE_EXECUTION_ACCESS_DENIED' -or $script:observedNextAction -ne 'CHECK_TARGET_EXECUTION_ACCESS' -or $result.ExitCode -ne $exitCode) { throw "Incorrect status for exit $exitCode" }
    $passed++
}
# Intermittent RPC then success: the real endpoint result is kept.
function Set-ResultFromAttempt { param($Attempt); $script:appliedAttempt=$Attempt; $result.LauncherStatus='ENDPOINT_COMPLETED' }
$script:observedStatus=''; $script:appliedAttempt=$null
$attemptResult=$recoveredAttempt; $transport=$recoveredTransport
$result=[ordered]@{ExitCode='';JobErrorMessage='';LauncherStatus='UNKNOWN'}
& $classify
if (-not $attemptResult.Completed -or $null -eq $script:appliedAttempt -or $script:appliedAttempt.Status -ne 'INTUNE_ENROLLMENT_PENDING_CONFIRMATION' -or $result.LauncherStatus -ne 'ENDPOINT_COMPLETED' -or $script:observedStatus -or $result.JobErrorMessage) { throw 'Endpoint result was not kept after RPC recovery' }
$passed++
# Continuous RPC failure beyond the window while the endpoint still runs: unconfirmed.
$script:observedStatus=''; $script:appliedAttempt=$null
$attemptResult=$lostAttempt; $transport=$lostTransport
$result=[ordered]@{ExitCode='';JobErrorMessage='';LauncherStatus='UNKNOWN'}
& $classify
if ($attemptResult.State -ne 'Running' -or $null -ne $script:appliedAttempt -or $script:observedStatus -ne 'REMOTE_EXECUTION_UNCONFIRMED' -or $script:observedNextAction -ne 'VERIFY_REMOTE_STATE_BEFORE_RELAUNCH') { throw 'Continuous RPC failure was not classified unconfirmed' }
$passed++
$catalogDefinition=$catalog.ByName['RemoteTaskRpcFailureWindowMinutes']
if ($null -eq $catalogDefinition -or [int]$catalogDefinition.Default -ne 30 -or [int]$catalogDefinition.Minimum -ne 0 -or [string]$catalogDefinition.Target -ne 'Technician') { throw 'RemoteTaskRpcFailureWindowMinutes catalog entry differs' }
$launcherText=[IO.File]::ReadAllText((Join-Path $root 'Scripts\SmartM365-Invoke-EndpointMigrationWithPsExec.ps1'))
$workerText=[IO.File]::ReadAllText($worker)
if ($launcherText -notmatch "RpcFailureWindowMinutes = \(Opt 'RemoteTaskRpcFailureWindowMinutes'\)" -or $workerText -notmatch "-RpcFailureWindowMinutes \(\[int\]\(Get-SemtPropertyValue -InputObject \`$contract.PsExec -Name 'RpcFailureWindowMinutes'" -or $workerText -notmatch '-RemoteDataShare \$remoteDataShare') { throw 'RPC failure window is not passed from the launcher to the transport' }
$passed++
Write-Host "Summary: Passed=$passed; Failed=0"

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD8Q9I2MvB62lGT
# fh7kA1OejgPrNRtDWirRvt2pJJOehKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCB68wZazcqzHyit/s38DAOR
# E/Bj3DsRK58UuGYakvXA0zANBgkqhkiG9w0BAQEFAASCAYCOGneVh8mpC32Y9ZEv
# biC8o0N2nWU8X8KWbnMqurWD8/CqmHQkBkzzUQ0KwuU3wx2rhQsXw5fD2YEuGlYn
# O/hCI/CuXZ9fZP8Z/88ZQuZS2iubD7HX/2IAlSBUEd1fXPHvoiv21/I0AcFIdYGQ
# iDN4nuV9PSPyITO28jYas/xDgTNaSgqiDfK4ok6PjuOiVHVpfozvM/S3LpKZtoyl
# c8V/5gzGbOtYg7pbBXEz2OmdxusO9fRKeqA3TrLRz/ykJjeLKsi/GyA6D7nHyioK
# gegkr8E8OygxP7S1qWvJVN/8QNnPaUcVbMHMAWCiPPEQZc0mEYFS7ptIIHX00FrE
# DH01IOYLrV6YN/DeYDa80r8QRJEwmIe8a0je2ErHGPnCJCctCft28fuZctOr984w
# OdQVh+YUP7/xBZieiiuolgKHMP0YX+qIhgMpUXi0nL8767xS8LZAUEynYlmBL/2V
# 92OxcQ3rx3kU1azQXOkyZpcObHPMzuLkaw1iLLtQls7a3pY=
# SIG # End signature block
