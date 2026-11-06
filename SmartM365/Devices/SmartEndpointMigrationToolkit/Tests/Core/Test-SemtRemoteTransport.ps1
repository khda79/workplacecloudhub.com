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
    # Replies without State/LastRunTime, as seen right after a dropped RPC connection.
    [int]$EmptyReplies = 0
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
        if ($this.EmptyReplies -gt 0) { $this.EmptyReplies--; return [pscustomobject]@{ State = $null; LastRunTime = $null } }
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
# Task of an unconfirmed attempt, as read by Remove-SemtUnstartedSystemTask.
class SemtCleanupTask {
    [SemtFakeDefinition]$Definition = [SemtFakeDefinition]::new()
    [int]$State = 3
    [datetime]$LastRunTime = [datetime]'1899-12-30'
    [bool]$Enabled = $true
}
class SemtCleanupFolder {
    [SemtCleanupTask]$Task = [SemtCleanupTask]::new()
    [bool]$Gone = $false
    # Simulates the trigger firing between the first read and the disable.
    [bool]$StartsWhenDisabled = $false
    [bool]$Deleted = $false
    [bool]$DisabledBeforeDelete = $false
    [object] GetTask([string]$name) {
        if ($this.Gone) { throw [System.Runtime.InteropServices.COMException]::new('The system cannot find the file specified. (Exception from HRESULT: 0x80070002)', -2147024894) }
        if ($this.StartsWhenDisabled -and -not $this.Task.Enabled) { $this.Task.State = 4; $this.Task.LastRunTime = [datetime]::Now }
        return $this.Task
    }
    [void] DeleteTask([string]$name,[int]$flags) { $this.DisabledBeforeDelete = -not $this.Task.Enabled; $this.Deleted = $true }
}
class SemtCleanupService {
    [SemtCleanupFolder]$Folder = [SemtCleanupFolder]::new()
    [void] Connect() { }
    [void] Connect([string]$target) { }
    [object] GetFolder([string]$path) { return $this.Folder }
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
foreach ($code in @(-2147023174,-2147023169,-2147023436,-2147024775,-2147417848,-2147024843,-2147024832)) {
    # 0x800706BA server unavailable, 0x800706BF call failed DNE, 0x800705B4 timeout, 0x80070079 semaphore timeout, 0x80010108 disconnected,
    # 0x80070035 network path not found and 0x80070040 network name deleted (target briefly off the network)
    $service=[SemtFakeService]::new(); $service.Folder.GetTaskFailures=2; $service.Folder.FailHResult=$code
    $scenario=Invoke-SemtRpcScenario -Service $service
    if ($scenario.Error -or -not $scenario.Output[0].Cleaned -or -not $service.Folder.Deleted) { throw ('Transient code 0x{0:X8} was not retried: {1}' -f $code,$scenario.Error) }
    $passed++
}
$service=[SemtFakeService]::new(); $service.Folder.EmptyReplies=2
$scenario=Invoke-SemtRpcScenario -Service $service
if ($scenario.Error -or -not $scenario.Output[0].Cleaned -or -not $service.Folder.Deleted -or @($scenario.Warnings | Where-Object { $_ -like 'SEMT_RPC_TRANSIENT: Step=GetTaskState; Code=0xINCOMPLETE*' }).Count -ne 2) { throw "An empty task reply was not retried: $($scenario.Error)" }
$passed++
$service=[SemtFakeService]::new(); $service.Folder.EmptyReplies=100000
$scenario=Invoke-SemtRpcScenario -Service $service -Window 15
if ($scenario.Error -notmatch '^SEMT_RPC_UNAVAILABLE: Step=GetTaskState; Code=0xINCOMPLETE' -or $service.Folder.Deleted -or -not $scenario.DirectoryKept) { throw "Endless empty task replies were not bounded by the window: $($scenario.Error)" }
$passed++
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
# Cleanup of the task of an unconfirmed attempt: deleted only when it never started.
$cleanupFn=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Remove-SemtUnstartedSystemTask'},$true)
if (-not $cleanupFn) { throw 'Remove-SemtUnstartedSystemTask not found' }
$cleanupAssignment=$cleanupFn.Find({param($n) $n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$operation'},$true)
$cleanupExpression=$cleanupAssignment.Right.Extent.Text
$cleanupOperation=[scriptblock]::Create($cleanupExpression.Substring(1,$cleanupExpression.Length-2))
$cleanupName='WCH-SEMT-22222222222222222222222222222222'
$cleanupRequest=@{Target='synthetic.invalid';Name=$cleanupName;Mode='ScheduledTask';Description=('SEMT temporary SYSTEM execution ' + $cleanupName)}
function New-SemtCleanupFake {
    $service=[SemtCleanupService]::new()
    $service.Folder.Task.Definition.RegistrationInfo.Description=('SEMT temporary SYSTEM execution ' + $cleanupName)
    return $service
}
$script:fake=New-SemtCleanupFake
$reply=& $cleanupOperation $cleanupRequest
if ($reply.Outcome -ne 'Deleted' -or -not $script:fake.Folder.Deleted -or -not $script:fake.Folder.DisabledBeforeDelete) { throw "A never-started task was not disabled then deleted: $($reply.Outcome)" }
$passed++
$script:fake=New-SemtCleanupFake; $script:fake.Folder.Task.State=4; $script:fake.Folder.Task.LastRunTime=[datetime]::Now
$reply=& $cleanupOperation $cleanupRequest
if ($reply.Outcome -ne 'Started' -or $script:fake.Folder.Deleted -or -not $script:fake.Folder.Task.Enabled) { throw 'A running task was touched' }
$passed++
$script:fake=New-SemtCleanupFake; $script:fake.Folder.Task.LastRunTime=[datetime]::Now.AddMinutes(-5)
$reply=& $cleanupOperation $cleanupRequest
if ($reply.Outcome -ne 'Started' -or $script:fake.Folder.Deleted -or -not $script:fake.Folder.Task.Enabled) { throw 'A task that already ran was touched' }
$passed++
$script:fake=New-SemtCleanupFake; $script:fake.Folder.StartsWhenDisabled=$true
$reply=& $cleanupOperation $cleanupRequest
if ($reply.Outcome -ne 'Started' -or $script:fake.Folder.Deleted) { throw 'A task that started while being disabled was deleted' }
$passed++
$script:fake=New-SemtCleanupFake; $script:fake.Folder.Gone=$true
$reply=& $cleanupOperation $cleanupRequest
if ($reply.Outcome -ne 'Absent' -or $script:fake.Folder.Deleted) { throw 'A missing task was not reported absent' }
$passed++
$script:fake=New-SemtCleanupFake; $script:fake.Folder.Task.Definition.RegistrationInfo.Description='Another task'
$reply=& $cleanupOperation $cleanupRequest
if ($reply.Outcome -ne 'Failed' -or $script:fake.Folder.Deleted -or -not $script:fake.Folder.Task.Enabled) { throw 'A task with another identity was touched' }
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
    # Cleanup of an unconfirmed attempt: the job still gives no reply (Wait-Job returns nothing).
    & $module {
        $script:CleanupJobs = 0
        function script:Start-Job { param($ArgumentList,$ScriptBlock); $script:CleanupJobs++; return [pscustomobject]@{State='Running'} }
    }
    $cleanupCorrelation='22222222222222222222222222222222'
    $noReply=Remove-SemtUnstartedSystemTask -Mode ScheduledTask -ConnectionTarget 'synthetic.invalid' -CorrelationId $cleanupCorrelation -TimeoutSeconds 10
    if ($noReply.Outcome -ne 'Failed' -or $noReply.Detail -notmatch 'kept') { throw 'A cleanup without reply was not reported failed with the task kept' }
    $passed++
    $toolkitRoot=Join-Path ([IO.Path]::GetTempPath()) ('SEMT-cleanup-' + [guid]::NewGuid().ToString('N'))
    $cleanupShare=Join-Path $toolkitRoot 'Prerequisites'
    try {
        [void][IO.Directory]::CreateDirectory((Join-Path $cleanupShare 'State\Results'))
        [void][IO.Directory]::CreateDirectory((Join-Path $toolkitRoot 'State'))
        [IO.File]::WriteAllText((Join-Path $toolkitRoot 'State\TargetOperation.json.txt'),('{"Campaign":"Prerequisites","CorrelationId":"' + $cleanupCorrelation + '"}'))
        & $module { $script:CleanupJobs = 0 }
        $leaseHeld=Remove-SemtUnstartedSystemTask -Mode WinRM -ConnectionTarget 'synthetic.invalid' -CorrelationId $cleanupCorrelation -RemoteDataShare $cleanupShare -TimeoutSeconds 10
        if ($leaseHeld.Outcome -ne 'Started' -or (& $module { $script:CleanupJobs }) -ne 0) { throw 'The target lease of the attempt did not keep the task untouched' }
        $passed++
        [IO.File]::WriteAllText((Join-Path $cleanupShare ('State\Results\{0}.json.txt' -f $cleanupCorrelation)),('{"CorrelationId":"' + $cleanupCorrelation + '","State":"Running"}'))
        $resultHeld=Remove-SemtUnstartedSystemTask -Mode ScheduledTask -ConnectionTarget 'synthetic.invalid' -CorrelationId $cleanupCorrelation -RemoteDataShare $cleanupShare -TimeoutSeconds 10
        if ($resultHeld.Outcome -ne 'Started' -or (& $module { $script:CleanupJobs }) -ne 0) { throw 'The attempt result did not keep the task untouched' }
        $passed++
        Remove-Item -LiteralPath (Join-Path $cleanupShare ('State\Results\{0}.json.txt' -f $cleanupCorrelation)) -Force
        [IO.File]::WriteAllText((Join-Path $toolkitRoot 'State\TargetOperation.json.txt'),'{"Campaign":"HybridJoin","CorrelationId":"33333333333333333333333333333333"}')
        & $module {
            function script:Wait-Job { param($Job,$Timeout); return $Job }
            function script:Receive-Job { param($Job); return [pscustomobject]@{Outcome='Deleted'; Detail='Task never started (State=3); disabled then deleted.'} }
        }
        $deleted=Remove-SemtUnstartedSystemTask -Mode ScheduledTask -ConnectionTarget 'synthetic.invalid' -CorrelationId $cleanupCorrelation -RemoteDataShare $cleanupShare -TimeoutSeconds 10
        if ($deleted.Outcome -ne 'Deleted' -or (& $module { $script:CleanupJobs }) -ne 1 -or $deleted.TaskName -ne ('WCH-SEMT-' + $cleanupCorrelation)) { throw 'Another campaign lease blocked the cleanup of a never-started task' }
        $passed++
    } finally { Remove-Item -LiteralPath $toolkitRoot -Recurse -Force -ErrorAction SilentlyContinue }
    & $module { function script:Wait-Job { param($Job,$Timeout); return $null } }
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
# Worker: an unconfirmed automatic Prerequisites run cleans its task only with a scheduled task transport.
$prerequisitesFn=$workerAst.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Invoke-WorkerPrerequisitesEndpoint'},$true)
if (-not $prerequisitesFn) { throw 'Invoke-WorkerPrerequisitesEndpoint not found' }
$prerequisitesText=$prerequisitesFn.Extent.Text
$unconfirmedAt=$prerequisitesText.IndexOf('if (-not $done.Completed)')
$guardAt=$prerequisitesText.IndexOf("if (`$remoteMode -eq 'PsExec')",[Math]::Max(0,$unconfirmedAt))
$psexecLogAt=$prerequisitesText.IndexOf('Write-SemtLogLine',[Math]::Max(0,$guardAt))
$elseAt=$prerequisitesText.IndexOf('else {',[Math]::Max(0,$guardAt))
$cleanupAt=$prerequisitesText.IndexOf('Remove-SemtUnstartedSystemTask',[Math]::Max(0,$elseAt))
$returnAt=$prerequisitesText.IndexOf("State = 'Unconfirmed'",[Math]::Max(0,$cleanupAt))
if ($unconfirmedAt -lt 0 -or $guardAt -lt 0 -or $psexecLogAt -lt 0 -or $psexecLogAt -gt $elseAt -or $cleanupAt -lt 0 -or $returnAt -lt 0 -or $prerequisitesText.Substring($returnAt,200) -notmatch '\$cleanupDetail') { throw 'Unconfirmed Prerequisites run does not clean its task before returning' }
if ($prerequisitesText -notmatch [regex]::Escape('Task cleanup not applicable (PsExec)')) { throw 'PsExec unconfirmed run does not log that the cleanup does not apply' }
$passed++
Write-Host "Summary: Passed=$passed; Failed=0"

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCkY2JNWbK+0fCK
# VN45JaKKyrujzLnfQg+ze/7Pk37FIaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIFJWslkyMae13/2dU3nAs8Km/RN6cWmQTgbKfV3BT9FZMA0GCSqG
# SIb3DQEBAQUABIIBgJStNsRob4eDri3B4NN86EqOZudQ85BaCGuanvAZK08cbm6q
# IyIZKt3AucgtguH0wKXpQF373HJhxCUisor2DJ29L7whkgOxfC1IuaETlsQfr7MD
# +oZcL14WCxdl4SMDzPmFluYMnpQn6AoMFRDj7WI+vwMvrr3FckV32wdZfKDro083
# oPOQis3AfNKAQxgvGH0YZiz2c4z1CEfIywjxdwVGT4xQrWh5evgLyrLHLHR5SIFV
# If3PvlAITzr8vGqAjctdA6dJ6cjjFLpLW7i253ROxLcQYc3gCiluF0CMQJ7UBLzs
# gHZo8vnJxq4F3S9+dsDBlllkunABWMMOOXaH23W3IzJRmRWB6mreVqwN8ThkqAb7
# C8O/tlf08qPVP5g52yslfn7YUWMD0wD9K/q58/16in3sNoFBY8zIt15GvA4AoZG8
# 3/iEle3d1Tb0j8DQ3+9vpnSW8m6wDnmNgGQgbTv3gd/Hc6roi7dTP8Y+vbooJDTa
# mKZKIhdCQCJIVRJSnaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDYyMTI5
# MTdaMC8GCSqGSIb3DQEJBDEiBCBP5P8Shc/gISfzY/XeaU7eA8S/E2gcfFowfCn7
# wrdOzTANBgkqhkiG9w0BAQEFAASCAgAHLkXkIjLZCqm2T4chWhBUCpRxGCN7IYnH
# zO5p7TRMrenuCqlb8ALwa6k/mId5a+6VtKcDwXaNhsxLnqJ+fF5eIDtPyvUAp2bN
# ct0eualMXTvvQJlIOmtIHZN7AIIBr5mSVqYVomeZ+3kmesGe7nuf3DBusXF03rGn
# pSooy2Fx2B6F8MO9+2e53pdAZpe6uc/eMdSEzePDWnnrpN3oGeWTDQtAxNLQPQgX
# D+dUQEtrHPl9+1MM/MM6uoXX/CatZJ26OhOOn4y9yfIVY3R5ATmGjYcK0GSl9qfA
# J2di6LYSHc7PHzZiG6hGQgFU9ZUvFyAj7mrAUhVwal91VLV0Zoi/mU7AtBynRCMq
# pIpC78rg3W25bu6VY6dPPkTwuq8Eqo5GTu8OLLWzDq7K6t7sSOkeETPQNtCLRauE
# FMN5avBWbu6t5WU/rS83vDTSM7ZhaL2oKpLBc2o/k+6TmiRJbOgsLoE9rafWdNOC
# AOyzpBMKJqOkL2U8qO2iHoXQFXAaTj3CD4F7kFA2ko+DPj5iQuFiZ9ayNd+n1fFN
# frYa5M49nStrdqznekmYK/WF1Nx6rtbwRkhqyH/JB8lA/0to6u7PaWttmAfEKzG2
# MX2w5svUR+sA877O3xy+EzdhAIDknFXzulVKiHZi4fB7akmEE9qOIHHcx9rc+5DZ
# 8hlbrX24ZQ==
# SIG # End signature block
